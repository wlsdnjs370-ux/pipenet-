# -*- coding: utf-8 -*-
"""link-prediction 연결복원 v2 — 학습데이터 생성 (T분기 1급 모델링).

v1 한계
=======
v1 은 끊긴 junction 을 'un-snap' 해 **끝점↔끝점**(tip-tip) 후보만 만들었다.
망단위 CV PR-AUC 0.81 로 피처/모델은 유효했으나, 실제 DXF 전이는 0건이었다.
원인(linkpred_diag.py): 실제 CAD 파탄의 참 연결은 가지배관 끝이 본관 **중간(edge)**
에 닿는 **T분기(끝점↔edge)** 가 지배적(LH306 63%, 대명동 39%) 인데, v1 은
끝점↔끝점만 봐서 구조적으로 T분기를 제안하지 못한다.

v2 전략
=======
한 망(SDF 정답)에 **두 종류의 합성손상**을 섞어 학습 분포를 실제에 맞춘다:

  (A) un-snap break  (tip-tip 양성)
      차수≥2 노드 N 의 인접 끝단들을 N 좌표 주변으로 흩뿌림 → 끊긴 끝단쌍.
      양성 = 같은 N 에서 떨어져 나온 끝단끼리.

  (B) T-tap 합성     (tip-edge 양성)
      차수3 노드 N 에서 가장 '직선연속'인 2배관(through-pair, cos≈-1)을
      하나의 본관 폴리라인 [far_a … N … far_b] 으로 **병합**하고, 남은 가지배관
      끝단을 N 근처로 떨어뜨림. 양성 = 가지 끝단 ↔ 병합 본관의 **내부(N 근처)**.
      → 끝점↔edge 정답을 합성한다.

후보 = tip→tip ∪ tip→edge 를 한 모델(`is_edge` 플래그 포함)에 통합. 트리가
is_edge 로 갈라 각 유형의 관련 피처만 쓰게 한다.

피처는 **기하-only** (DXF 엔 구경 없음 — [[pipe-material-not-in-dxf]]).
"""
from __future__ import annotations

import math
import random
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from kfp_sdf_converter import parse_sdf  # noqa: E402

FEATURE_NAMES = [
    "is_edge",            # 0=tip-tip, 1=tip-edge
    "gap", "gap_norm",
    "tip_align",          # tip.out_dir 가 표적을 향하는 각 (0=정면)
    "tt_partner_align",   # tip-tip: 상대 끝단 out_dir 가 이쪽을 향하는 각 / edge: 0
    "tt_cos_dirs",        # tip-tip: 두 끝단 방향 코사인(직선연속≈-1) / edge: 0
    "te_proj_interior",   # edge: min(p,1-p) 본관 내부도 ∈[0,.5] / tip-tip: 0
    "te_entry_angle",     # edge: 가지 진입각 vs 본관 국소방향 [0,90] / tip-tip: 0
]


# ── 손상 후 구조 ─────────────────────────────────────────────────────────────
@dataclass
class _Seg:
    """본관(또는 일반 배관) 한 가닥. 병합되면 폴리라인이 된다."""
    id: int
    pts: list                      # [(x,y), ...]  (>=2)
    absorbed: set = field(default_factory=set)   # 병합으로 내부에 흡수된 node_id 들


@dataclass
class _Tip:
    """끊긴(loose) 끝단."""
    id: int
    coord: tuple
    out_dir: tuple                 # 배관 내부→끝단 방향 (단위)
    node_id: str                   # 손상 전 원래 노드 (라벨용)
    seg_id: int                    # 자기 자신이 속한 seg (자기연결 배제용)


# ── 기하 헬퍼 ────────────────────────────────────────────────────────────────
def _unit(dx, dy):
    n = math.hypot(dx, dy)
    return (dx / n, dy / n) if n > 1e-12 else (0.0, 0.0)


def _angle_deg(u, v):
    c = max(-1.0, min(1.0, u[0] * v[0] + u[1] * v[1]))
    return math.degrees(math.acos(c))


def _acute(a):
    return a if a <= 90.0 else 180.0 - a


def _node_xy(net, nid):
    n = net.nodes.get(nid)
    return (n.x, n.y) if n is not None else (0.0, 0.0)


def _median(vals):
    s = sorted(v for v in vals if v > 0)
    return s[len(s) // 2] if s else 1.0


def _pt_polyline(p, pts):
    """점 p → 폴리라인 최근접. (dist, nearest, sub_dir, global_param) 반환.

    global_param ∈ [0,1] = 폴리라인 길이 기준 위치(끝=0/1, 중간=내부).
    """
    total = 0.0
    cum = [0.0]
    for i in range(len(pts) - 1):
        total += math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1])
        cum.append(total)
    best = (float("inf"), pts[0], (0.0, 0.0), 0.0)
    for i in range(len(pts) - 1):
        ax, ay = pts[i]
        bx, by = pts[i + 1]
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        if L2 <= 1e-18:
            continue
        t = ((p[0] - ax) * dx + (p[1] - ay) * dy) / L2
        t = max(0.0, min(1.0, t))
        cx, cy = ax + t * dx, ay + t * dy
        d = math.hypot(p[0] - cx, p[1] - cy)
        if d < best[0]:
            gp = (cum[i] + t * math.sqrt(L2)) / total if total > 0 else 0.0
            best = (d, (cx, cy), _unit(dx, dy), gp)
    return best


# ── 합성손상 ─────────────────────────────────────────────────────────────────
def build_damaged(net, *, rng, break_prob=0.5, ttap_prob=0.6,
                  jitter_lo=0.2, jitter_hi=1.5,
                  tap_lo=0.0, tap_hi=1.0):
    """깨끗한 SDF → (_Seg 목록, _Tip 목록, med_edge).

    차수3 노드는 ttap_prob 로 T-tap 합성(through-pair 병합 + 가지 떨굼),
    그 외 차수≥2 노드는 break_prob 로 un-snap break(끝단 흩뿌림).
    차수1(헤드·수원 등 진짜 단말)은 손상 안 함.

    jitter 범위 분리: un-snap break 끝단은 [jitter_lo,jitter_hi](≈1세그먼트로 벌어짐),
    T-tap 가지 끝단은 [tap_lo,tap_hi](본관에 거의 얹히는 실측 분포 — gap≈0 포함).
    """
    # 노드별 인접 끝단 (pid, side) + 좌표/길이
    incident: dict[str, list[tuple[str, str]]] = defaultdict(list)
    for pid, p in net.pipes.items():
        incident[p.start].append((pid, "s"))
        incident[p.end].append((pid, "e"))

    edge_len = {}
    for pid, p in net.pipes.items():
        a = _node_xy(net, p.start)
        b = _node_xy(net, p.end)
        edge_len[pid] = math.hypot(b[0] - a[0], b[1] - a[1])
    med_edge = _median(edge_len.values())

    def pipe_pts(pid):
        p = net.pipes[pid]
        return [_node_xy(net, p.start)] + [tuple(w) for w in (p.waypoints or [])] \
            + [_node_xy(net, p.end)]

    def side_dir(pid, side):
        """배관 내부→해당 끝단 단위방향 (앵커 = 인접 vertex)."""
        pl = pipe_pts(pid)
        if side == "s":
            tip, anc = pl[0], pl[1]
        else:
            tip, anc = pl[-1], pl[-2]
        return _unit(tip[0] - anc[0], tip[1] - anc[1])

    # ── 1) T-tap 노드 선정 (through-pair 결정, 배관 소비) ──
    consumed: set[str] = set()        # through-pair 로 병합된 pid
    ttap_plan = {}                    # node_id -> (through_a, through_b, branch_pid, branch_side)
    deg = {nid: len(eps) for nid, eps in incident.items()}
    for nid, eps in incident.items():
        if deg[nid] != 3 or rng.random() >= ttap_prob:
            continue
        if any(pid in consumed for pid, _ in eps):
            continue
        # 각 인접 배관의 N 으로부터 바깥 방향(내부→N 의 반대 = N→내부)
        dirs = []
        for pid, side in eps:
            d = side_dir(pid, side)           # 내부→끝단(=N)
            dirs.append((pid, side, (-d[0], -d[1])))   # N→배관내부
        # through-pair = 두 배관이 가장 직선(서로 반대방향, cos≈-1)
        best, pair = 2.0, None
        for i in range(3):
            for j in range(i + 1, 3):
                c = dirs[i][2][0] * dirs[j][2][0] + dirs[i][2][1] * dirs[j][2][1]
                if c < best:
                    best, pair = c, (i, j)
        if pair is None or best > -0.5:       # 충분히 직선인 통과쌍이 없으면 skip
            continue
        i, j = pair
        k = 3 - i - j
        a_pid, a_side, _ = dirs[i]
        b_pid, b_side, _ = dirs[j]
        br_pid, br_side, _ = dirs[k]
        consumed.add(a_pid)
        consumed.add(b_pid)
        ttap_plan[nid] = (a_pid, a_side, b_pid, b_side, br_pid, br_side)

    # ── 2) break 노드 선정 (T-tap·소비된 배관 제외) ──
    break_nodes: set[str] = set()
    for nid, eps in incident.items():
        if nid in ttap_plan or deg[nid] < 2:
            continue
        if rng.random() < break_prob:
            break_nodes.add(nid)

    # ── 3) 끝단 jitter 좌표 결정 ──
    def jitter(base, local_scale, lo, hi):
        mag = rng.uniform(lo, hi) * max(local_scale, med_edge * 0.1)
        ang = rng.uniform(0, 2 * math.pi)
        return (base[0] + mag * math.cos(ang), base[1] + mag * math.sin(ang))

    # 각 (pid, side) 의 손상 후 좌표 — 기본은 원좌표
    ep_coord: dict[tuple, tuple] = {}
    tips_meta: list[tuple] = []   # (coord, out_dir, node_id, pid, side)

    for pid, p in net.pipes.items():
        for side, nid in (("s", p.start), ("e", p.end)):
            base = _node_xy(net, nid)
            if nid in break_nodes and pid not in consumed:
                c = jitter(base, edge_len.get(pid) or med_edge, jitter_lo, jitter_hi)
                ep_coord[(pid, side)] = c
                tips_meta.append((c, side_dir(pid, side), nid, pid, side))
            else:
                ep_coord[(pid, side)] = base

    # ── 4) 세그먼트 구성 ──
    segs: list[_Seg] = []
    pipe_seg: dict[str, int] = {}     # 일반 배관 pid -> seg.id
    sid = 0

    # 4a) 병합(through-pair) 세그먼트 + 가지 끝단
    for nid, (a_pid, a_side, b_pid, b_side, br_pid, br_side) in ttap_plan.items():
        N = _node_xy(net, nid)
        # A: far_a … N
        a_pl = pipe_pts(a_pid)
        a_far = a_pl[::-1] if a_side == "s" else a_pl   # N 이 끝에 오도록
        # a_far 는 [far_a, ..., N]
        # B: N … far_b
        b_pl = pipe_pts(b_pid)
        b_from = b_pl if b_side == "s" else b_pl[::-1]   # [N, ..., far_b]
        merged = a_far[:-1] + [N] + b_from[1:]
        segs.append(_Seg(id=sid, pts=merged, absorbed={nid}))
        sid += 1
        # 가지 끝단을 N 근처로 떨굼 (본관에 거의 얹히는 분포)
        c = jitter(N, edge_len.get(br_pid) or med_edge, tap_lo, tap_hi)
        ep_coord[(br_pid, br_side)] = c
        tips_meta.append((c, side_dir(br_pid, br_side), nid, br_pid, br_side))

    # 4b) 나머지 일반 배관 (소비 안 된 것) — 손상 후 좌표로 폴리라인 구성
    for pid, p in net.pipes.items():
        if pid in consumed:
            continue
        mids = [tuple(w) for w in (p.waypoints or [])]
        pl = [ep_coord[(pid, "s")]] + mids + [ep_coord[(pid, "e")]]
        segs.append(_Seg(id=sid, pts=pl, absorbed=set()))
        pipe_seg[pid] = sid
        sid += 1

    # ── 5) Tip 객체화 (seg_id 부여) ──
    tips: list[_Tip] = []
    tid = 0
    for coord, odir, nid, pid, side in tips_meta:
        seg_id = pipe_seg.get(pid, -1)        # 병합 가지는 자기 seg 없음(-1)
        tips.append(_Tip(id=tid, coord=coord, out_dir=odir,
                         node_id=nid, seg_id=seg_id))
        tid += 1
    return segs, tips, med_edge


# ── 후보 + 피처 ──────────────────────────────────────────────────────────────
def candidate_rows(segs, tips, med_edge, *, search_factor=2.5,
                   interior_lo=0.04):
    """tip→tip ∪ tip→edge 후보 + 피처/라벨.

    rows: (feat_dict, label, info_dict). info = {kind, a, b}
      kind: 'tt' | 'te',  a: tip 좌표,  b: 상대 끝단/본관 최근접점.
    """
    R = search_factor * med_edge
    R_sq = R * R
    rows: list[tuple[dict, int, dict]] = []

    # tip 공간격자 (tip-tip 용)
    cell = max(R, 1e-9)
    tgrid = defaultdict(list)
    for t in tips:
        tgrid[(int(t.coord[0] // cell), int(t.coord[1] // cell))].append(t.id)
    tip_by_id = {t.id: t for t in tips}

    # ── tip → tip ──
    seen = set()
    for ti in tips:
        kx, ky = int(ti.coord[0] // cell), int(ti.coord[1] // cell)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for jid in tgrid.get((kx + dx, ky + dy), ()):
                    if jid == ti.id:
                        continue
                    key = (min(ti.id, jid), max(ti.id, jid))
                    if key in seen:
                        continue
                    tj = tip_by_id[jid]
                    if ti.seg_id != -1 and ti.seg_id == tj.seg_id:
                        continue   # 같은 배관 양끝
                    gap2 = (tj.coord[0] - ti.coord[0]) ** 2 + (tj.coord[1] - ti.coord[1]) ** 2
                    if gap2 == 0 or gap2 > R_sq:
                        continue
                    seen.add(key)
                    label = 1 if ti.node_id == tj.node_id else 0
                    rows.append((_feat_tt(ti, tj, med_edge), label,
                                 {"kind": "tt", "a": ti.coord, "b": tj.coord}))

    # ── tip → edge ──
    for t in tips:
        for seg in segs:
            if seg.id == t.seg_id:
                continue
            dist, near, sub_dir, gp = _pt_polyline(t.coord, seg.pts)
            if dist > R or dist <= 1e-9:
                continue
            if min(gp, 1.0 - gp) < interior_lo:
                continue       # 본관 끝부분 투영은 tip-tip 영역 — 제외
            label = 1 if t.node_id in seg.absorbed else 0
            rows.append((_feat_te(t, dist, sub_dir, gp, med_edge), label,
                         {"kind": "te", "a": t.coord, "b": near}))
    return rows


def _feat_tt(ti: _Tip, tj: _Tip, med_edge) -> dict:
    gap = math.hypot(tj.coord[0] - ti.coord[0], tj.coord[1] - ti.coord[1])
    ab = _unit(tj.coord[0] - ti.coord[0], tj.coord[1] - ti.coord[1])
    ba = (-ab[0], -ab[1])
    return {
        "is_edge": 0.0,
        "gap": gap,
        "gap_norm": gap / med_edge if med_edge > 0 else gap,
        "tip_align": _angle_deg(ti.out_dir, ab),
        "tt_partner_align": _angle_deg(tj.out_dir, ba),
        "tt_cos_dirs": ti.out_dir[0] * tj.out_dir[0] + ti.out_dir[1] * tj.out_dir[1],
        "te_proj_interior": 0.0,
        "te_entry_angle": 0.0,
    }


def _feat_te(t: _Tip, dist, sub_dir, gp, med_edge) -> dict:
    # 표적 = 본관 위 최근접점; tip→표적 방향
    # (gp 로 정확 좌표 복원 대신, tip_align 은 진입각의 보완정보 te_entry_angle 로 흡수)
    return {
        "is_edge": 1.0,
        "gap": dist,
        "gap_norm": dist / med_edge if med_edge > 0 else dist,
        "tip_align": 0.0,     # tip-edge 에선 진입각(te_entry_angle)이 주신호
        "tt_partner_align": 0.0,
        "tt_cos_dirs": 0.0,
        "te_proj_interior": min(gp, 1.0 - gp),
        "te_entry_angle": _acute(_angle_deg(t.out_dir, sub_dir)),
    }


def corpus_files(mode="remote"):
    """학습 코퍼스 파일 목록.

    mode="remote": 스프링클러 REMOTE 43건 (v2 baseline).
    mode="all"   : 답안 SDF 전체(수리계산 참고용 도서/**.sdf). 위상 다양성↑.
                   파싱/규모 필터는 build_matrix 쪽에서 처리.
    """
    import validate_sdf as vs
    if mode == "remote":
        return vs._remote_answer_files()
    import glob as _glob
    pat = str(_ROOT / "수리계산 참고용 도서" / "**" / "*.sdf")
    return sorted(_glob.glob(pat, recursive=True))


def dataset_from_net(net, *, seed=0, **kw) -> list[tuple[dict, int]]:
    rng = random.Random(seed)
    segs, tips, med = build_damaged(
        net, rng=rng,
        break_prob=kw.get("break_prob", 0.5),
        ttap_prob=kw.get("ttap_prob", 0.6),
        jitter_lo=kw.get("jitter_lo", 0.2),
        jitter_hi=kw.get("jitter_hi", 1.5),
        tap_lo=kw.get("tap_lo", 0.0),
        tap_hi=kw.get("tap_hi", 1.0))
    rows = candidate_rows(segs, tips, med,
                          search_factor=kw.get("search_factor", 2.5))
    return [(f, lab) for (f, lab, _info) in rows]


if __name__ == "__main__":
    import validate_sdf as vs
    files = vs._remote_answer_files()
    print(f"REMOTE 답안 {len(files)}건")
    f0 = files[0]
    net = parse_sdf(f0)
    ds = dataset_from_net(net, seed=1)
    pos_tt = sum(1 for f, l in ds if l == 1 and f["is_edge"] == 0)
    pos_te = sum(1 for f, l in ds if l == 1 and f["is_edge"] == 1)
    n_tt = sum(1 for f, _ in ds if f["is_edge"] == 0)
    n_te = sum(1 for f, _ in ds if f["is_edge"] == 1)
    print(f"표본: {Path(f0).name}")
    print(f"  tip-tip 후보 {n_tt} (양성 {pos_tt}) · tip-edge 후보 {n_te} (양성 {pos_te})")

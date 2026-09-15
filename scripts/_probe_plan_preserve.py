# -*- coding: utf-8 -*-
"""[아이소·평면보존 §2] 손질 정본 G 가 수리계산까지 «그대로» 갔는가 — 먼저 잰다.

지시서 `ModuleF_아이소_평면보존_지시서.md` · 그림 `docs/images/iso_plan_preserve/`.

■ 판정식 (§1-3)

    G′ = G_K ∪ ⋃_h T_h        G_K 의 노드·간선·XY 는 그대로, T_h 만 «덧붙는다»

    P1  G_K 의 모든 노드 v 에 π(v′) = p(v) 인 G′ 노드가 있다
    P2  G_K 의 모든 간선 (u,v) 가 π(G′) 에서 같은 XY 의 사슬로 남아 있다
    P3  π(G′) 에 G_K 에도 T_h 에도 속하지 않는 노드·간선이 없다
    H1  T_h 가 1-2 의 자리에 선다 — 상향·상하향은 H, 하향은 T (팔 없으면 H)
    H2  노즐 노드의 π = p(H)   (상하향 D 만 p(H) + ③·u)
    H3  토막 길이 = 입력값 · 재질 = 템플릿 복사 · 후렉시블은 하향② · 상하향④
    I1  아이소 좌표 = x′=(x−y)cos30 · y′=(x+y)sin30 + (z−z_ref)·lift

  훼손의 네 얼굴: (a) 뿌리가 H/T 가 아니다 · (b) XY 가 옮겨졌다 ·
                 (c) 평면에 없던 것이 생겼다 · (d) 평면에 있던 것이 사라졌다

■ 무엇을 도는가 (§2-1)

  `/design/build` 가 부르는 **바로 그 함수를 같은 순서로** 부른다. 새 경로를
  만들지 않는다 — `expand_worst` 가 하는 세 줄(restrict → planar → vertical)을
  같은 함수로 펼쳐 부를 뿐이다. 그래야 **단계마다** 같은 판정을 돌려 «어디서
  처음 깨지는지» 를 볼 수 있다(planar 직후 · vertical 직후 · tables 직후).

    python scripts/_probe_plan_preserve.py [--key 저장본] [--k 30]
                                           [--zone x0 y0 x1 y1] [--source Z1]
"""
from __future__ import annotations

import argparse
import math
import os
import sys
from collections import defaultdict, deque
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DM_KEY = "1. 입력도면 대명동 단위세대 평면도"

# [가지치기·부속판정 §2] planar 가 로그로만 내는 값(`끝배관1단보호`)을 읽으려고
# 화면 출력을 그대로 흘리면서 사본을 모은다. `main()` 이 채운다.
_LOG = None


# ─────────────────────────────────────────────── 자·거리
def _d(a, b):
    return math.hypot(float(a[0]) - float(b[0]), float(a[1]) - float(b[1]))


def _hist(vals, cuts=(5.0, 50.0, 150.0)):
    """오차 목록 → (최대, {문턱: 초과 개수})."""
    out = {c: sum(1 for v in vals if v > c) for c in cuts}
    return (max(vals) if vals else 0.0), out


class Plan:
    """kfp(m) ↔ 평면(mm) 환산 — `build_planar_graph` 의 xform 을 되돌린다.

        mx = (x_mm − minx)/1000 + 1      ⇒   x_mm = mx·1000 + (minx − 1000)
    """

    def __init__(self, origin_mm):
        o = origin_mm or (0.0, 0.0)
        self.ox = float(o[0]) - 1000.0
        self.oy = float(o[1]) - 1000.0

    def mm(self, coords):
        return (float(coords[0]) * 1000.0 + self.ox,
                float(coords[1]) * 1000.0 + self.oy)


def nodes_of(kfp):
    return dict((kfp or {}).get("nodes_meta_runtime") or {})


def pipes_of(kfp):
    return dict((kfp or {}).get("pipe_data") or {})


def _ends(pr):
    return pr.get("start") or pr.get("from"), pr.get("end") or pr.get("to")


# ─────────────────────────────────────────────── P1 · P2 (자는 옆 모듈에)
#
#   ★자를 한 번 틀렸다. 「G_K 노드마다 같은 자리의 G′ **노드**가 있나」로 재면
#     `build_planar_graph` 의 **일직선 중간 노드 병합**이 전부 「사라짐」으로
#     세어진다(실측 대명동 520 → 286 · 끊김 300/519). 지시서 §1-3 P2 는 그
#     병합을 허용한다 — 「세분·병합만 허용」. 그래서 «점·선분 집합» 으로 잰다.
from _pp_checks import (PlanSet, check_P1, check_P2,      # noqa: E402
                        check_P3)


# ─────────────────────────────────────────────── 세로 토막
def vertical_stacks(kfp, plan, tol=1.0):
    """세로 배관 = 평면에서 길이 0 인 배관. {평면XY(반올림): [(pid, dz)]}"""
    nd, pr = nodes_of(kfp), pipes_of(kfp)
    out = defaultdict(list)
    for pid, p in pr.items():
        s, e = _ends(p)
        a = (nd.get(s) or {}).get("coords")
        b = (nd.get(e) or {}).get("coords")
        if not a or not b:
            continue
        pa, pb = plan.mm(a), plan.mm(b)
        if _d(pa, pb) <= tol:
            dz = float(b[2]) - float(a[2]) if len(a) > 2 and len(b) > 2 else 0.0
            out[(round(pa[0], 1), round(pa[1], 1))].append((pid, dz, s, e))
    return out


def head_rows(kfp, plan, node_head_kinds):
    """kfp 의 헤드 노드 — (nid, 평면XY, z, 종류, 붙은 배관 수)."""
    nd, pr = nodes_of(kfp), pipes_of(kfp)
    deg = defaultdict(int)
    for p in pr.values():
        s, e = _ends(p)
        deg[s] += 1
        deg[e] += 1
    out = []
    for nid, m in nd.items():
        if str((m or {}).get("type_id") or "").lower() != "head" and \
           str((m or {}).get("type") or "") != "Head":
            continue
        c = m.get("coords") or (0, 0, 0)
        out.append({"nid": nid, "xy": plan.mm(c),
                    "z": float(c[2]) if len(c) > 2 else 0.0,
                    "kind": (node_head_kinds or {}).get(nid),
                    "deg": deg.get(nid, 0)})
    return out


# ─────────────────────────────────────────────── 본체
def run(key, k, zones=None, source=None, top=8):
    # ★서버가 부팅 때 못박는 **쓰기 루트**를 그대로 쓴다. `import_write_root()`
    #   의 기본값은 상대경로 `docs/import` 라, 저장소 루트에서 돌리면 찍은스펙·
    #   표시캐시를 엉뚱한 폴더에서 찾는다(실측: DWG 옛 경로로 떨어져 MISS).
    #   서버 모듈을 한 번 들이면 `set_write_root` 가 이미 불려 있다 —
    #   /design/build 가 보는 것과 같은 판을 보게 된다.
    from routes.module_f.common import _boot
    _boot()                       # 라우트가 첫 요청에서 부르는 바로 그 줄

    from services.cad_import.edit.io import load_edits, open_board
    from services.cad_import.design.restrict import (
        attachable_heads, restrict_to_worst, apply_vertical)
    from services.cad_import.design.worst import worst_k_heads
    from services.cad_import.convert.planar import build_planar_graph
    from services.cad_import.design.tables import build_design_tables
    from services.cad_import.dto import default_dto, dto_to_convert_kwargs

    board = open_board(key)
    load_edits(board)

    # ── 급수원 — 저장본에 없으면 손질 화면과 **같은 길로** 하나 찍는다.
    #   (`EditSession.click` 이 모듈 F 의 `/edit/click` 이 부르는 그 줄이다.
    #    board 에 직접 쓰면 손질이 하는 스냅·검증을 건너뛰게 된다.)
    if not board.sources:
        from services.cad_import.edit.board import body_seg_groups
        from services.cad_import.edit.session import MODE_SOURCE, EditSession
        groups = body_seg_groups(board.pts, board.edges, board.bodies())
        segs = sorted((g[0] for g in groups), key=len, reverse=True)[0]
        es = EditSession(board, key=key)
        es.set_mode(MODE_SOURCE)
        pa, pb = segs[0]              # 한 선분 = (점, 점)
        es.click((float(pa[0]) + float(pb[0])) / 2.0,
                 (float(pa[1]) + float(pb[1])) / 2.0, 2000)
        print(f"  [준비] 급수원을 찍었습니다 — {len(board.sources)}곳")

    # ★payload 는 **급수원을 찍은 뒤에** 뜬다 — `EditSession.convert_payload`
    #   와 같은 내용이다. 먼저 뜨면 급수원이 안 실려 물길 필터가 못 돈다.
    payload = dict(board.payload())
    payload["key"] = key
    payload["pts"] = [list(p) for p in board.pts]
    payload["edges"] = [list(e) for e in board.edges]
    payload["hcov"] = [list(d) for d in board.disks]
    payload["ups"] = [list(u) for u in board.ups]
    payload["disk_kinds"] = list(board.disk_kinds)
    payload["edge_len_mm"] = dict(getattr(board, "edge_len_mm", None) or {})

    # ── 후보: 영역 제한만 (손질과 같은 자)
    only = None
    if zones:
        only = {i for i, dsk in enumerate(board.disks)
                if any(z[0] <= float(dsk[0]) <= z[2]
                       and z[1] <= float(dsk[1]) <= z[3] for z in zones)}
    src_index = 0 if board.sources else None
    sel = worst_k_heads(board.pts, board.edges, board.hnodes, board.sources,
                        k=k, only_heads=only, source_index=src_index,
                        head_xy=board.disks)
    picked = [int(h) for h in (sel.get("heads") or ())]
    if not picked:
        print("★최불리가 비었다 — 급수원·이음을 확인하세요")
        return 1

    # ── /design/build 와 같은 순서 (expand_worst 의 세 줄을 펼쳐 부른다)
    probe = attachable_heads(payload, selected_source=source, key=key)
    cand = set(picked) & set(probe.get("wet") or ())
    worst = worst_k_heads(board.pts, board.edges, board.hnodes, board.sources,
                          k=k, only_heads=cand, head_xy=board.disks)
    limited = restrict_to_worst(payload, board, worst)
    # ★`expand_worst` 와 같은 규칙 — 스냅을 끄되 회랑이 쪼개지면 되돌린다.
    snap_off = True
    built = build_planar_graph(
        key, write=False, selected_source=source,
        pts=limited.get("pts"), edges=limited.get("edges"),
        hcov=limited.get("hcov"), ups=limited.get("ups"),
        head_kinds=limited.get("head_kinds"),
        user_sources=limited.get("sources"), ho=limited.get("ho"),
        edge_len_mm=limited.get("edge_len_mm"),
        grid_snap=snap_off,
        # ★[가지치기 §3-1] `expand_worst` 가 회랑에 거는 그 값. 프로브가 이걸
        #   빠뜨리면 **제품과 다른 회랑**을 재게 된다 — 잔류 스텁이 남은 판을
        #   재고 「제품이 내는 망」이라 부르는 셈이다.
        keep_head_stub=False)
    # ★`expand_worst` 와 같은 되돌림 — 스냅을 끄니 회랑이 쪼개지면 켠 채로.
    def _ok(b):
        kf = (b or {}).get("kfp")
        if not (b or {}).get("ok") or kf is None:
            return False
        nd, pp = nodes_of(kf), pipes_of(kf)
        hn = sum(1 for m in nd.values()
                 if str((m or {}).get("type_id") or "") == "head")
        if hn < len(worst.get("heads") or ()):
            return False
        ad = defaultdict(list)
        for pr in pp.values():
            a, z = _ends(pr)
            if a and z:
                ad[a].append(z)
                ad[z].append(a)
        if not ad:
            return False
        st = [next(iter(ad))]
        sn = set(st)
        while st:
            c = st.pop()
            for x in ad.get(c, ()):
                if x not in sn:
                    sn.add(x)
                    st.append(x)
        return len(sn) >= len(nd)

    if not _ok(built):
        print("  [준비] 스냅을 끄니 회랑이 쪼개집니다 — 켠 채로 갑니다")
        snap_off = False
        built = build_planar_graph(
            key, write=False, selected_source=source,
            pts=limited.get("pts"), edges=limited.get("edges"),
            hcov=limited.get("hcov"), ups=limited.get("ups"),
            head_kinds=limited.get("head_kinds"),
            user_sources=limited.get("sources"), ho=limited.get("ho"),
            edge_len_mm=limited.get("edge_len_mm"), grid_snap=False,
            keep_head_stub=False)
    if not built.get("ok"):
        print(f"★평면 전개 실패 — {built.get('error')}")
        return 1
    flat = built["kfp"]
    plan = Plan(built.get("origin_mm"))
    kw = dto_to_convert_kwargs(default_dto())
    raised, verr = apply_vertical(payload, built, convert_kwargs=kw)
    if raised is None:
        print(f"★세로 처리 실패 — {verr}")
        return 1

    # ★[회랑 사슬좌표 §3-1] `expand_worst` 가 여기서 부르는 그 줄. 프로브도
    #   같은 자리에서 같은 함수를 불러야 «제품이 내는 망» 을 재는 것이 된다.
    from services.cad_import.design.chain import chain_coords
    chain = chain_coords(raised,
                         edge_ref=built.get("edge_ref") or {},
                         node_ref=built.get("node_ref") or {},
                         board_pts=board.pts,
                         origin_mm=built.get("origin_mm"),
                         declared_pipes=built.get("declared_pipes") or (),
                         corridor_edges=worst.get("edges") or ())
    # ★C1 은 **합치기 전에** 잰다. 합치기가 가운데 절점을 지우므로 뒤에서
    #   재면 `legs` 가 가리키는 노드가 사라져 엉뚱한 수가 나온다
    #   (내가 한 번 39,317mm 를 보고 위반이라 셌다 — 자가 낡은 것이었다).
    def _xyz0(nid):
        c = ((raised.get("nodes_meta_runtime") or {}).get(nid)
             or {}).get("coords") or (0, 0, 0)
        return (float(c[0]), float(c[1]),
                float(c[2]) if len(c) > 2 else 0.0)

    c1_err = 0.0
    _acc = {chain["root"]: _xyz0(chain["root"])}
    _ch = defaultdict(list)
    for _pid, _lg in (chain.get("legs") or {}).items():
        _ch[_lg["par"]].append((_lg["chi"], _lg["vec"]))
    _q = deque([chain["root"]])
    while _q:
        _cur = _q.popleft()
        for _nx, _v in _ch.get(_cur, ()):
            if _nx in _acc:
                continue
            _p0 = _acc[_cur]
            _acc[_nx] = (_p0[0] + _v[0], _p0[1] + _v[1], _p0[2] + _v[2])
            c1_err = max(c1_err, math.dist(_acc[_nx], _xyz0(_nx)) * 1000.0)
            _q.append(_nx)

    from services.cad_import.design.chain import merge_straight_runs
    mrg = merge_straight_runs(raised,
                              edge_ref=built.get("edge_ref") or {},
                              node_ref=built.get("node_ref") or {},
                              declared_pipes=built.get("declared_pipes") or ())
    if mrg.get("merged"):
        built["edge_ref"] = mrg["edge_ref"]

    gk_nodes = {int(v) for v in (worst.get("nodes") or ())}
    gk_edges = {(int(a), int(b)) for a, b in (worst.get("edges") or ())}
    node_ref = built.get("node_ref") or {}
    edge_ref = built.get("edge_ref") or {}

    print(f"\n■ 평면 보존 계측 · {key} · K={k}"
          f" · 영역 {len(zones or ())}곳 · 급수원 {source or '자동'}")
    print(f"\n  G   노드 {len(board.pts)} · 간선 {len(board.edges)}"
          f" · 헤드 {len(board.disks)}")
    print(f"  G_K corridor 노드 {len(gk_nodes)} · 간선 {len(gk_edges)}"
          f" · 헤드 {len(picked)}")
    print(f"  G′  평면직후 노드 {len(nodes_of(flat))}"
          f" · 배관 {len(pipes_of(flat))}"
          f"  →  세로후 노드 {len(nodes_of(raised))}"
          f" · 배관 {len(pipes_of(raised))}")

    # ── 단계별 같은 판정 (어디서 처음 깨지나)
    stage = {}
    for label, net in (("평면 직후", flat), ("세로 직후", raised)):
        ps = PlanSet(net, plan, nodes_of, pipes_of, _ends)
        p1 = check_P1(gk_nodes, board.pts, ps)
        p2 = check_P2(gk_edges, board.pts, ps)
        stage[label] = (p1, p2)
        print(f"\n  ── {label}")
        print(f"  [P1] G_K 노드 {p1['n']} · 노드로 남음 {p1['on_node']}"
              f" · 간선 위로(중간노드 병합 · 허용) {len(p1['merged'])}"
              f" · 격자 스냅(≤35mm · §6-1) {len(p1['snapped'])}"
              f" · ★사라짐 {len(p1['lost'])}")
        print(f"       XY 오차 최대 {p1['max']:.1f}mm"
              f" · >5mm {p1['over'][5.0]}"
              f" · >50mm {p1['over'][50.0]}"
              f" · >150mm {p1['over'][150.0]}")
        for v, q, dd in p1["lost"][:top]:
            print(f"        ★사라짐 board {v} ({q[0]:.0f}, {q[1]:.0f})"
                  f" · π(G′) 까지 {dd:.0f}mm")
        for v, q, dd in p1["snapped"][:3]:
            print(f"        ·스냅 board {v} ({q[0]:.0f}, {q[1]:.0f})"
                  f" · {dd}mm")
        print(f"  [P2] G_K 간선 {p2['n']} · 덮임 {p2['kept']}"
              f" · ★끊김 {len(p2['broken'])}"
              f" · 5~150mm 이탈 {len(p2['strayed'])}"
              f" · 최대 이탈 {p2['max']:.1f}mm")
        for u, v, w in p2["broken"][:top]:
            print(f"        ★끊김 ({u},{v})"
                  f" ({board.pts[u][0]:.0f}, {board.pts[u][1]:.0f})–"
                  f"({board.pts[v][0]:.0f}, {board.pts[v][1]:.0f})"
                  f" · 최대 {w}mm")
        for u, v, w in p2["strayed"][:top]:
            print(f"        ·이탈 ({u},{v}) 최대 {w}mm")

    p3 = check_P3(flat, raised, node_ref, nodes_of)
    print(f"\n  [P3] G′ 노드 {p3['total']} = G_K 대응 {p3['gk']}"
          f" + T_h {p3['th']} + 미분류 {len(p3['unknown'])}")
    nd_f = nodes_of(flat)
    for nid in p3["unknown"][:top]:
        c = (nd_f.get(nid) or {}).get("coords") or (0, 0, 0)
        q = plan.mm(c)
        print(f"        ★미분류 {nid} ({q[0]:.0f}, {q[1]:.0f})")

    # ── H1 · H2 — 헤드 자리와 뿌리
    nhk = built.get("node_head_kinds") or {}
    heads = head_rows(raised, plan, nhk)
    stacks = vertical_stacks(raised, plan)
    disks = board.disks
    h2_bad, h1_bad = [], []
    kind_n = defaultdict(int)
    for h in heads:
        # 이 헤드가 평면의 어느 헤드인가 — 가장 가까운 disk
        bi, bd = None, 1e18
        for i in picked:
            dd = _d(h["xy"], (float(disks[i][0]), float(disks[i][1])))
            if dd < bd:
                bi, bd = i, dd
        kind = h["kind"] or (board.disk_kinds[bi] if bi is not None
                             and bi < len(board.disk_kinds) else "?")
        kind_n[str(kind)] += 1
        # H2 — 노즐 XY = p(H) (상하향 아래 헤드 D 만 옆으로 벗어난다)
        if bd > 5.0 and "상하향" not in str(kind):
            h2_bad.append((h["nid"], h["xy"], bi, round(bd, 1), str(kind)))
        # H1 — 뿌리: 이 헤드 자리에 세로 토막이 서 있나
        key_xy = (round(h["xy"][0], 1), round(h["xy"][1], 1))
        near_stack = any(_d(k2, h["xy"]) <= 5.0 for k2 in stacks)
        if not near_stack:
            h1_bad.append((h["nid"], h["xy"], str(kind)))
    print(f"\n  [H1] 헤드 {len(heads)}개 · 종류 "
          + " · ".join(f"{k2} {v}" for k2, v in sorted(kind_n.items())))
    print(f"       제 자리에 세로 토막 있음 {len(heads) - len(h1_bad)}"
          f" · 없음 {len(h1_bad)}")

    # ★[그림 4 · 읽기 A] 하향식 ① 은 **팔이 갈라지는 티 T** 에 선다.
    #   p(H) 의 세로(②)만 보면 읽기 A 와 읽기 B 를 못 가른다 — 둘 다 H 에
    #   세로가 있다. 가르는 것은 «T 에도 세로가 서 있나» 다.
    nd_v, pr_v = nodes_of(raised), pipes_of(raised)
    adj_h = defaultdict(list)
    for pid, p in pr_v.items():
        s2, e2 = _ends(p)
        a2 = (nd_v.get(s2) or {}).get("coords")
        b2 = (nd_v.get(e2) or {}).get("coords")
        if not a2 or not b2:
            continue
        if _d(plan.mm(a2), plan.mm(b2)) > 1.0:        # 가로 배관만
            adj_h[s2].append(e2)
            adj_h[e2].append(s2)
    #   ★단, **팔이 없는** 하향식은 ① 없이 H 에서 ② 만 내린다(§1-2 · 현 규칙).
    #     그러니 「팔 있음」 만 골라 세야 한다 — 안 가르면 규칙대로인 것을
    #     훼손으로 센다. 팔 있음 = p(H) 에 선 노드의 가로 배관이 **하나**
    #     (가지 끝에서 팔 하나로 매달린 모양). 둘 이상이면 가지관 위 통과점이라
    #     팔이 없다(§2-2 가 통과관을 쪼개 붙인 헤드가 이 부류다).
    #   ★«티가 어디냐» 를 내가 다시 걸어 찾지 않는다. 꺾인 팔은 티가 여러
    #     마디 건너 있어서, 바로 옆 마디를 티로 보면 틀린다(실측: 헤드에서
    #     42mm 떨어진 팔 위의 점을 티로 집었다). 그 걸음은 `_pendant_arm_to_tee`
    #     가 이미 갖고 있으므로, 여기서는 **세로가 선 자리를 그냥 센다** —
    #     헤드 자리가 아닌 세로가 곧 ① 의 자리다.
    pend = [h for h in heads if "하향" in str(h["kind"] or "")]
    at_head = sum(1 for k3 in stacks
                  if any(_d(k3, h["xy"]) <= 5.0 for h in heads))
    print(f"       세로가 선 자리 {len(stacks)}곳 — 헤드 자리 {at_head}"
          f" · 헤드 아닌 자리 {len(stacks) - at_head}"
          f"  (하향식 {len(pend)}개 · 헤드 아닌 세로 = ① 의 자리)")
    for nid, q, kind in h1_bad[:top]:
        print(f"        ★{nid} ({q[0]:.0f}, {q[1]:.0f}) {kind}"
              f" — 세로 토막이 이 자리에 없다")
    print(f"  [H2] 노즐 XY = p(H) {len(heads) - len(h2_bad)}/{len(heads)}")
    # ★어긋난 것이 «격자 스냅» 인지 **단정하지 않고 못박는다.**
    #   planar 의 xform 은 (x−minx)/1000+1 를 GRID_M=0.05 로 반올림한다.
    #   그러니 p(H) 를 같은 식으로 스냅한 값이 실제 π(노즐)과 같으면, 그것이
    #   원인이다 — 다른 무엇이 옮긴 것이 아니다(§6-1 오너 결정 자리).
    GRID = 50.0

    def snapped_mm(q):
        return (round((q[0] - plan.ox) / GRID) * GRID + plan.ox,
                round((q[1] - plan.oy) / GRID) * GRID + plan.oy)

    n_grid = sum(1 for _n, q, bi, _d2, _k in h2_bad
                 if _d(q, snapped_mm((float(disks[bi][0]),
                                      float(disks[bi][1])))) <= 0.2)
    print(f"       그중 «격자 스냅(50mm)로 설명되는 것» {n_grid}/{len(h2_bad)}"
          f"  ← 설명되면 §6-1 오너 결정, 아니면 다른 원인")
    # 설명 안 되는 것을 **먼저** 낸다 — 그것이 볼 자리다.
    def _explained(rec):
        _n, q2, bi2, _d3, _k2 = rec
        return _d(q2, snapped_mm((float(disks[bi2][0]),
                                  float(disks[bi2][1])))) <= 0.2
    for nid, q, bi, dd, kind in sorted(h2_bad, key=_explained)[:top]:
        sq = snapped_mm((float(disks[bi][0]), float(disks[bi][1])))
        print(f"        ★{nid} ({q[0]:.0f}, {q[1]:.0f}) ≠ disk {bi}"
              f" ({disks[bi][0]:.0f}, {disks[bi][1]:.0f}) · {dd}mm · {kind}"
              f" · 격자로 스냅하면 ({sq[0]:.0f}, {sq[1]:.0f})"
              + ("  = 같다(격자 스냅)"
                 if _d(q, sq) <= 0.2 else "  ★★다르다 — 격자가 아니다"))

    # ── H3 — 토막 길이가 입력값인가 · 후렉시블 자리
    want = {round(float(kw.get(n2) or 0), 3)
            for n2 in ("pendant_rise_m", "pendant_drop_m", "upright_rise_m",
                       "combo_1_m", "combo_2_m", "combo_3_m", "combo_up_m")
            if kw.get(n2) is not None}
    pr_r = pipes_of(raised)
    vlen, vbad = 0, []
    flexc = kw.get("flex_c")
    n_flex = 0
    for xy2, lst in stacks.items():
        for pid, dz, s, e in lst:
            vlen += 1
            L = round(float((pr_r.get(pid) or {}).get("length_m") or 0), 3)
            if want and L not in want:
                vbad.append((pid, L))
            if flexc is not None and \
                    float((pr_r.get(pid) or {}).get("C") or 0) == float(flexc):
                n_flex += 1
    print(f"\n  [H3] 세로 토막 {vlen}개 · 길이 = 입력값"
          f" {vlen - len(vbad)}/{vlen}"
          f" · 후렉시블 {n_flex}개 (입력값 {sorted(want)})")
    nd_r = nodes_of(raised)
    for pid, L in vbad[:top]:
        pp = pr_r.get(pid) or {}
        s2, e2 = _ends(pp)
        za = (nd_r.get(s2) or {}).get("coords") or (0, 0, 0)
        zb = (nd_r.get(e2) or {}).get("coords") or (0, 0, 0)
        q = plan.mm(za)
        # 세로 토막이 헤드 자리인가 아닌가 — 헤드가 아니면 «가지 상승·급수 라이저»
        near_h = min((_d(q, h["xy"]) for h in heads), default=1e18)
        print(f"        ★{pid} 길이 {L} m — 입력값에 없는 값"
              f" · ({q[0]:.0f}, {q[1]:.0f}) z {za[2]:.2f}→{zb[2]:.2f}"
              f" · 가장 가까운 헤드 {near_h:.0f}mm"
              + ("  (헤드 자리 아님 — 가지 상승/급수 라이저로 보인다)"
                 if near_h > 50 else ""))

    # ══════════════ [회랑 사슬좌표 §2] 바꾸기 전에 재는 네 가지
    from _pp_chain import (head_slip_bound, measure_far, measure_merge,
                           measure_merge_actual,
                           measure_tree, measure_turns, path_len)

    turns = measure_turns(gk_edges, board.pts)
    dd8 = turns["dist"]
    print(f"\n  [회전각] 회랑 평면 배관 {turns['n']}개 · 8방향과의 편차:"
          f" 0°(±0.5) {dd8[0.5]} · ≤5° {dd8[5.0]}"
          f" · ≤10° {dd8[10.0]} · ≤22.5° {dd8[22.5]}"
          f" · 최대 {turns['max']:.2f}°")
    for r in turns["rows"][:min(10, top)]:
        if r["dev"] <= 0.5:
            break
        print(f"        ·({r['u']},{r['v']}) L {r['L'] / 1000:.2f}m"
              f" · 편차 {r['dev']:.2f}° → {r['dir']:.0f}°"
              f" · 밀림 {r['slip']:.1f}mm")
    slip = head_slip_bound(worst, board.pts, turns)
    hslip = [slip["acc"].get(int(n), 0.0)
             for n in (worst.get("nodes") or ())]
    print(f"           헤드 이탈 상한 Σ L·sin(편차) — 최대"
          f" {max(hslip, default=0.0):.1f}mm"
          f" (회랑 전체 최대 {slip['max']:.1f}mm)")

    # ★«결과» 를 본다 — 스냅이 켜졌는지 꺼졌는지 믿지 않는다.
    mg = measure_merge_actual(gk_nodes, board.pts,
                              PlanSet(raised, plan, nodes_of, pipes_of, _ends),
                              built.get("origin_mm") or (0, 0))
    print(f"  [병합]   G_K 노드 {len(gk_nodes)} (결과 망에서 직접 셈)"
          f" · 같은 좌표(중복 절점 · 정상) {mg['n_same']}"
          f" + ★떨어진 좌표 {mg['n_apart']}"
          + ("   ← §6-2" if mg["n_apart"] else ""))
    for cell, (vs, far) in list(mg["apart"].items())[:top]:
        pp = [f"board {v}({board.pts[v][0]:.0f},{board.pts[v][1]:.0f})"
              for v in vs]
        print(f"        ★한 칸 {cell} ← " + " · ".join(pp)
              + f"  (서로 {far:.0f}mm)")

    tr = measure_tree(raised, nodes_of, pipes_of, _ends)
    print(f"  [나무]   회랑 |V| {tr['V']} · |E| {tr['E']}"
          f" · 연결성분 {tr['comp']}"
          f" · |E|=|V|−1 이고 성분 1 인가: "
          + ("예" if tr["tree"] else "★아니오"))

    from services.cad_import.design.anchor import require_anchor
    root_nid = require_anchor(nodes_of(raised), what="사슬 계측")
    fctx = measure_far(worst, board.pts, raised, root_nid,
                       nodes_of, pipes_of, _ends)
    anchor_h = min(heads, key=lambda h: -h["z"]) if heads else None
    # 앵커 = 손질이 고른 최원 헤드. 그 자리의 kfp 헤드를 좌표로 찾는다.
    wh = int(worst.get("worst_head") or -1)
    tgt = None
    if 0 <= wh < len(disks):
        pw = (float(disks[wh][0]), float(disks[wh][1]))
        tgt = min(heads, key=lambda h: _d(h["xy"], pw))["nid"] if heads else None
    kfp_far = path_len(fctx, tgt, plan) if tgt else None
    print(f"  [최원길이] 손질 far_m(board 거리) {fctx['far_m']:.3f}m"
          f" · 경로 재계산(board) {fctx['board_m']:.3f}m"
          f" · 표의 평면 길이 합 "
          + (f"{kfp_far:.3f}m · 차 {abs(kfp_far - fctx['far_m']) * 1000:.0f}mm"
             if kfp_far is not None else "(앵커 헤드를 못 찾음)"))

    # ══════════════ [회랑 사슬좌표 §4-1] C1~C4 — 앞 지시서 P1·H2 를 **대체**
    nd_c, pr_c = nodes_of(raised), pipes_of(raised)
    decl = {str(p) for p in (built.get("declared_pipes") or ())}
    legs = chain.get("legs") or {}

    def _xyz(nid):
        c = (nd_c.get(nid) or {}).get("coords") or (0, 0, 0)
        return (float(c[0]), float(c[1]),
                float(c[2]) if len(c) > 2 else 0.0)

    # C2 — length_m 이 board 유클리드(평면)·입력값(세로)이고 좌표 거리와 같은가
    c2_bad, c2_err = [], 0.0
    for pid, pr in pr_c.items():
        if str(pid) in decl:
            continue
        s3, e3 = _ends(pr)
        d3 = math.dist(_xyz(s3), _xyz(e3))
        L = float(pr.get("length_m") or 0.0)
        er = abs(d3 - L) * 1000.0
        c2_err = max(c2_err, er)
        if er > 1.0:
            c2_bad.append((pid, round(L, 4), round(d3, 4), round(er, 2)))
    # C3 — 1:1 · 끊김 0 · 미분류 0
    _p1s, _p2s = stage["세로 직후"]
    c3_ok = (mg["n_apart"] == 0 and not _p2s["broken"]
             and not p3["unknown"])
    # C4 — 8 방향 · 회전각 ≤ 22.5°
    c4_ok = (chain.get("dev_over_22_5", 0) == 0)
    wdiff = (abs(kfp_far - fctx["far_m"]) * 1000.0
             if kfp_far is not None else None)

    kk = chain.get("kinds") or {}
    print(f"\n  [사슬] 노드 {chain.get('nodes')} · 평면 {kk.get('plane', 0)}"
          f" · 세로 {kk.get('vert', 0)} · 엔진가로 {kk.get('engine', 0)}"
          f" · 길이 덮음 {chain.get('len_overwritten')}"
          f" · 기울어진 평면 {chain.get('skew', 0)}"
          f" · 닿지 못한 노드 {len(chain.get('unreached') or ())}")
    print(f"  [C1] 사슬 닫힘 — 최대 오차 {c1_err:.3f}mm"
          f"  {'OK' if c1_err <= 1.0 else '★위반'}")
    print(f"  [C2] 길이 = 좌표 거리 — 어긋난 배관 {len(c2_bad)}개"
          f" · 최대 {c2_err:.3f}mm"
          f"  {'OK' if not c2_bad else '★위반'}"
          f"  (선언 배관 {len(decl)}개 제외)")
    for pid, L, d3, er in c2_bad[:top]:
        print(f"        ★{pid} length_m {L} ≠ 좌표 {d3} · {er}mm")
    print(f"  [C3] 1:1 — 병합(떨어진 좌표) {mg['n_apart']}"
          f" · 끊김 {len(_p2s['broken'])} · 미분류 {len(p3['unknown'])}"
          f"  {'OK' if c3_ok else '★위반'}")
    print(f"  [C4] 8방향 — 최대 편차 {chain.get('dev_max', 0):.2f}°"
          f" · >22.5° {chain.get('dev_over_22_5', 0)}"
          f" · >10° {chain.get('dev_over_10', 0)}"
          f"  {'OK' if c4_ok else '★위반'}")
    print(f"  [W ] 손질 far_m ↔ 표 평면 길이 합 — 차 "
          + (f"{wdiff:.3f}mm  {'OK' if wdiff <= 1.0 else '★위반'}"
             if wdiff is not None else "(못 잼)"))
    print(f"  [이탈] 헤드·노드 실제 이탈 최대 {chain.get('slip_max_mm', 0):.1f}mm"
          f" (예상 상한 {max(hslip, default=0.0):.1f}mm"
          f" · 뿌리 스냅 {chain.get('root_snap_mm') or 0:.1f}mm) — 참고")

    # ── 표 · I1
    # ★[가지치기·부속판정 §3-2] 제품이 넘기는 그 차수를 프로브도 넘긴다.
    #   안 넘기면 표가 종전 규칙으로 서서, 프로브가 «제품이 내는 부속» 이
    #   아니라 «옛 부속» 을 재게 된다.
    from services.cad_import.design.restrict import corridor_topology
    _topo = corridor_topology(
        {"pts": limited.get("pts"), "edges": limited.get("edges")},
        {"node_ref": built.get("node_ref") or {}, "edge_ref": edge_ref,
         "origin_mm": built.get("origin_mm")}, raised)
    tbl = build_design_tables(raised, worst, edge_ref, [],
                              board_pts=board.pts,
                              origin_mm=built.get("origin_mm"),
                              tree_loads=None,
                              phys=_topo["phys"],
                              interior_junctions=_topo["interior_junctions"],
                              node_head_kinds=nhk)
    noz = [str(z.get("in")) for z in tbl.nozzles]
    at = {str(n.get("label")): (float(n.get("x") or 0), float(n.get("y") or 0))
          for n in tbl.nodes}
    print(f"\n  [표] 노드 {len(tbl.nodes)} · 배관 {len(tbl.pipes)}"
          f" · 노즐 {len(noz)} (K={len(picked)})")

    from services.cad_import.design.sdf_post import bake_isometric
    import copy as _copy
    t2 = _copy.deepcopy(tbl)
    before = {str(n.get("label")): (float(n.get("x") or 0),
                                    float(n.get("y") or 0),
                                    float(n.get("elevation") or 0))
              for n in t2.nodes}
    parent = {}
    for prw in tbl.pipes:
        a, z2 = str(prw.get("in")), str(prw.get("out"))
        if z2 in noz:
            parent[z2] = a
        if a in noz:
            parent[a] = z2
    info = bake_isometric(t2, head_nodes=noz, head_parent=parent)
    C30, S30 = math.cos(math.radians(30.0)), math.sin(math.radians(30.0))
    elevs = [v[2] for v in before.values()]
    e_ref = (min(elevs) + max(elevs)) / 2.0 if elevs else 0.0
    xs = [v[0] for v in before.values()]
    ys = [v[1] for v in before.values()]
    diag = math.hypot(max(xs) - min(xs), max(ys) - min(ys)) if xs else 0.0
    er = (max(elevs) - min(elevs)) if elevs else 0.0
    lift = (diag * 0.5 / er) if er > 0 else 0.0
    ok_i, bad_i, mx_i = 0, 0, 0.0
    for n in t2.nodes:
        lab = str(n.get("label"))
        if lab in noz:
            continue                      # 헤드는 부모 위에 세운다(별도 규칙)
        x0, y0, e0 = before[lab]
        ex = (x0 - y0) * C30
        ey = (x0 + y0) * S30 + (e0 - e_ref) * lift
        dd = math.hypot(float(n.get("x") or 0) - ex,
                        float(n.get("y") or 0) - ey)
        mx_i = max(mx_i, dd)
        if dd <= 1e-6:
            ok_i += 1
        else:
            bad_i += 1
    print(f"  [I1] 아이소 재계산 일치 {ok_i}/{ok_i + bad_i}"
          f" (최대 오차 {mx_i:.6f}) · 세운 헤드 {info.get('heads')}"
          f" · 관말 아님 {info.get('not_terminal')}")

    # ── [가지치기·부속판정 §2] 바꾸기 전에 재는 여섯 줄
    #
    #   ★**아무것도 안 고친 판**에서 잰다. 새 규칙은 옆 모듈이 «판정만» 돌려
    #     보고 파일을 쓰지 않는다 — 지금 값과 새 값을 같은 판에서 맞대야
    #     「몇 개가 바뀌나」가 뜻을 갖는다.
    import _pp_prune as pp
    try:
        prune_m = pp.report(key=key, k=k, limited=limited, built=built,
                            raised=raised, tbl=tbl, plan=plan, flat=flat,
                            stub_log=pp.stub_count_from_log(_LOG.text()),
                            payload=payload, board=board, worst=worst)
    except Exception as exc:                        # noqa: BLE001
        # 계측이 못 서도 앞의 판정식은 그대로 보고한다 — 자 하나가 고장 났다고
        # 나머지 수치를 잃으면 안 된다.
        print(f"\n  [§2 계측] ★못 쟀습니다 — {type(exc).__name__}: {exc}")
        prune_m = None

    # ── 훼손 분류
    p1r, p2r = stage["세로 직후"]
    # ★[사슬 §1-4] 앞 지시서의 P1·H2(좌표 = 평면)는 C1·C3 으로 **대체**됐다.
    #   사슬은 좌표를 «만들어» 쓰므로 평면과 어긋나는 것이 **규칙**이다
    #   (오너가 허용 · 상한 Σ L·sin(편차)). 그래서 (b) 는 여기서 세지 않고
    #   [이탈] 로 «참고» 보고한다 — 안 그러면 규칙대로인 것을 훼손으로 센다.
    a_ = len(h1_bad)
    b_ = 0
    c_ = len(p3["unknown"])
    d_ = len(p2r["broken"])
    print(f"\n  [훼손 분류] (a) 뿌리 {a_} · (b) XY 옮김 — C1·C3 으로 대체(참고)"
          f" · (c) 없던 것 {c_} · (d) 사라짐 {d_}")
    bad = a_ + b_ + c_ + d_
    fails = []
    if c1_err > 1.0:
        fails.append("C1")
    if c2_bad:
        fails.append("C2")
    if not c3_ok:
        fails.append("C3")
    if not c4_ok:
        fails.append("C4")
    if wdiff is None or wdiff > 1.0:
        fails.append("W")
    if bad:
        fails.append(f"H1/P3/P2({bad})")
    print("\n  " + ("★C1·C2·C3·C4·W 전부 선다 — 사슬 좌표가 맞다"
                    if not fails else f"★★안 서는 것: {' · '.join(fails)}"))
    return 0 if not fails else 2


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=DM_KEY)
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--zone", nargs=4, type=float, action="append")
    ap.add_argument("--source", default=None)
    ap.add_argument("--top", type=int, default=8)
    a = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    # ★[§2 잔류] `끝배관1단보호` 는 planar 가 **로그로만** 낸다. 반환에 키를
    #   더하면 지시서 §5 의 「planar.py diff = 키워드 통과뿐」이 깨지므로,
    #   화면에 그대로 흘리면서 사본을 모아 그 줄을 읽는다.
    global _LOG
    import _pp_prune as pp
    _LOG = pp.Tee(sys.stdout)
    real, sys.stdout = sys.stdout, _LOG
    try:
        return run(a.key, a.k, zones=a.zone, source=a.source, top=a.top)
    finally:
        sys.stdout = real


if __name__ == "__main__":
    raise SystemExit(main())

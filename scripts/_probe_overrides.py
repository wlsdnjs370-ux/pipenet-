# -*- coding: utf-8 -*-
"""[요소속성 수정카드 §2] 바꾸기 전에 «지금» 을 표로 — 짐작하지 않는다.

지시서 `ModuleF_요소속성_수정카드_지시서.md` §2 · 그림 14·15.

  낸다:
    [속성 표]  속성 × 지금 고칠 수 있나 · 어느 창 · 어느 키 · 표/‥sdf 의 어느 함수
    [키]       kfp 배관·노드 → 안정 키를 만들 수 있는 것 몇/몇 · 못 만드는 것의 갈래
    [생존]     K 30→20 · 영역 변경 뒤 살아남는 키 몇/몇
    [통합]     merge/preview 의 view 가 회랑 요소를 어떻게 가리키는가

    python scripts/_probe_overrides.py [--key 저장본] [--k 30] [--k2 20]
"""
from __future__ import annotations

import argparse
import math
import os
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DM_KEY = "1. 입력도면 대명동 단위세대 평면도"


def _ends(pr):
    return pr.get("start") or pr.get("from"), pr.get("end") or pr.get("to")


def _xyz(nd, nid):
    c = (nd.get(str(nid)) or {}).get("coords") or (0.0, 0.0, 0.0)
    return (float(c[0]), float(c[1]),
            float(c[2]) if len(c) > 2 else 0.0)


def build(board, payload, key, k, zones=None, source=None):
    """/design/build 가 밟는 그 길 — 표까지."""
    from services.cad_import.design.restrict import select_and_expand
    from services.cad_import.design.worst import worst_k_heads
    from services.cad_import.design.tables import build_design_tables

    only = None
    if zones:
        only = {i for i, d in enumerate(board.disks)
                if any(z[0] <= float(d[0]) <= z[2] and z[1] <= float(d[1]) <= z[3]
                       for z in zones)}
    sel = worst_k_heads(board.pts, board.edges, board.hnodes, board.sources,
                        k=k, only_heads=only,
                        source_index=0 if board.sources else None,
                        head_xy=board.disks)
    picked = [int(h) for h in (sel.get("heads") or ())]
    got = select_and_expand(payload, board, k=k, only_heads=set(picked),
                            selected_source=source, key=key)
    if not got.get("ok"):
        return None, got.get("error")
    tbl = build_design_tables(got["kfp"], got["worst"], got["edge_ref"], [],
                              board_pts=board.pts,
                              origin_mm=got.get("origin_mm"),
                              tree_loads=got.get("tree_loads"),
                              node_head_kinds=got.get("node_head_kinds"))
    return {"got": got, "tbl": tbl, "picked": picked}, None


def classify(got, board):
    """kfp 배관·노드마다 «안정 키를 만들 수 있나» 와 못 만들면 어느 갈래인가."""
    kfp = got["kfp"]
    nd = kfp.get("nodes_meta_runtime") or {}
    pr = kfp.get("pipe_data") or {}
    eref = {str(k): v for k, v in (got.get("edge_ref") or {}).items()
            if v is not None}
    nref = {str(k): int(v) for k, v in (got.get("node_ref") or {}).items()}
    o = got.get("origin_mm") or (0.0, 0.0)
    ox, oy = float(o[0]) - 1000.0, float(o[1]) - 1000.0

    def mm(nid):
        p = _xyz(nd, nid)
        return (p[0] * 1000.0 + ox, p[1] * 1000.0 + oy, p[2])

    heads = {}
    for nid, m in nd.items():
        if str((m or {}).get("type_id") or "") == "head":
            q = mm(nid)
            bi, bd = None, 1e18
            for i, d in enumerate(board.disks):
                dd = math.hypot(q[0] - float(d[0]), q[1] - float(d[1]))
                if dd < bd:
                    bi, bd = i, dd
            heads[str(nid)] = (bi if bd <= 60.0 else None, bd)

    pipe_keyed, pipe_un = {}, defaultdict(list)
    for pid, p in pr.items():
        pid = str(pid)
        s, e = _ends(p)
        if pid in eref:
            a, b = int(eref[pid][0]), int(eref[pid][1])
            pipe_keyed[pid] = ("pipe", min(a, b), max(a, b))
            continue
        qa, qb = mm(s), mm(e)
        flat = math.hypot(qa[0] - qb[0], qa[1] - qb[1])
        if flat <= 1.0:
            # 세로 토막 — 뿌리가 헤드인가 (그러면 («vert», disk, 역할) 이 가능)
            root = None
            for nid in (s, e):
                if str(nid) in heads and heads[str(nid)][0] is not None:
                    root = heads[str(nid)][0]
            pipe_un["A. 세로 토막 · 뿌리=헤드" if root is not None
                    else "B. 세로 토막 · 뿌리=헤드 아님(가지 상승·알람밸브)"
                    ].append(pid)
        else:
            pipe_un["C. 엔진이 만든 가로 토막(상하향 ③)"].append(pid)

    node_keyed = {nid: ("node", nref[nid]) for nid in nd if nid in nref}
    node_un = defaultdict(list)
    for nid in nd:
        if nid in node_keyed:
            continue
        if nid in heads and heads[nid][0] is not None:
            node_un["헤드(노즐) — disk 번호로 키 가능"].append(nid)
        else:
            node_un["엔진이 만든 절점(세로 중간·알람밸브)"].append(nid)
    return {"nd": nd, "pr": pr, "heads": heads,
            "pipe_keyed": pipe_keyed, "pipe_un": pipe_un,
            "node_keyed": node_keyed, "node_un": node_un}


def vert_roles(cls, nd):
    """세로 토막이 뿌리 헤드마다 몇 개인가 — §7-2 (역할 구별 가능한가)."""
    per = defaultdict(list)
    for kind, pids in cls["pipe_un"].items():
        if not kind.startswith("A. 세로 토막"):
            continue
        for pid in pids:
            s, e = _ends(cls["pr"][pid])
            root = None
            for nid in (s, e):
                h = cls["heads"].get(str(nid))
                if h and h[0] is not None:
                    root = h[0]
            per[root].append(pid)
    hist = defaultdict(int)
    for root, pids in per.items():
        hist[len(pids)] += 1
    return per, dict(sorted(hist.items()))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=DM_KEY)
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--k2", type=int, default=20)
    a = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.io import load_edits, open_board

    board = open_board(a.key)
    load_edits(board)
    if not board.sources:
        from services.cad_import.edit.board import body_seg_groups
        from services.cad_import.edit.session import MODE_SOURCE, EditSession
        g = sorted((x[0] for x in body_seg_groups(
            board.pts, board.edges, board.bodies())), key=len, reverse=True)[0]
        es0 = EditSession(board, key=a.key)
        es0.set_mode(MODE_SOURCE)
        pa, pb = g[0]
        es0.click((float(pa[0]) + float(pb[0])) / 2.0,
                  (float(pa[1]) + float(pb[1])) / 2.0, 2000)
    payload = dict(board.payload())
    payload["key"] = a.key
    payload["pts"] = [list(p) for p in board.pts]
    payload["edges"] = [list(e) for e in board.edges]
    payload["hcov"] = [list(d) for d in board.disks]
    payload["ups"] = [list(u) for u in board.ups]
    payload["disk_kinds"] = list(board.disk_kinds)
    payload["edge_len_mm"] = dict(getattr(board, "edge_len_mm", None) or {})

    print(f"\n■ 요소속성 수정카드 §2 계측 · {a.key}")

    r1, err = build(board, payload, a.key, a.k)
    if err:
        print(f"★빌드 실패 — {err}")
        return 1
    c1 = classify(r1["got"], board)
    npipe, nnode = len(c1["pr"]), len(c1["nd"])
    pk, nk = len(c1["pipe_keyed"]), len(c1["node_keyed"])
    print(f"\n[키]  K={a.k} · 표 배관 {npipe} · 절점 {nnode}")
    print(f"      배관 안정 키 가능 {pk}/{npipe}"
          f"  ({pk / npipe * 100:.0f}%)")
    for kind, pids in sorted(c1["pipe_un"].items(), key=lambda t: -len(t[1])):
        print(f"        · 불가 {len(pids):>3}개 — {kind}")
    print(f"      절점 안정 키 가능 {nk}/{nnode}"
          f"  ({nk / nnode * 100:.0f}%)")
    for kind, nids in sorted(c1["node_un"].items(), key=lambda t: -len(t[1])):
        print(f"        · {len(nids):>3}개 — {kind}")

    per, hist = vert_roles(c1, c1["nd"])
    print(f"\n[세로 토막 역할]  뿌리 헤드 {len(per)}개 · 헤드당 토막 수 분포"
          f" {hist}")
    print(f"      → 헤드당 1개면 («vert», disk, 역할) 로 바로 구별,"
          f" 2개 이상이면 z 순서로 갈라야 한다 (§7-2)")

    # ── 생존: K 를 바꿔 다시 빌드
    r2, err2 = build(board, payload, a.key, a.k2)
    if err2:
        print(f"\n[생존] K={a.k2} 빌드 실패 — {err2}")
    else:
        c2 = classify(r2["got"], board)
        keys1 = set(c1["pipe_keyed"].values())
        keys2 = set(c2["pipe_keyed"].values())
        alive = len(keys1 & keys2)
        print(f"\n[생존]  K {a.k}→{a.k2} · 배관 안정 키 {len(keys1)} 중"
              f" {alive}개 생존 ({alive / max(1, len(keys1)) * 100:.0f}%)")
        print(f"        (회랑이 줄면 그만큼 자리가 사라진다 — 사라진 것은"
              f" «적용 못 한 수정» 으로 올라가야 한다)")
        # kfp 이름이 얼마나 옮겨 다니는지 — 안정 키가 필요한 이유
        lab1 = {str(p.get("label") or pid) for pid, p in c1["pr"].items()}
        moved = 0
        for pid, skey in c1["pipe_keyed"].items():
            pid2 = next((q for q, kk in c2["pipe_keyed"].items()
                         if kk == skey), None)
            if pid2 is not None and pid2 != pid:
                moved += 1
        print(f"        ★kfp 이름(P#)으로 키를 삼으면: 같은 자리인데 이름이"
              f" 바뀐 배관 {moved}개 — 그만큼 죽는다")

    # ── 표에 실린 속성 · .sdf 로 가는 자리
    tbl = r1["tbl"]
    pf = sorted((tbl.pipes[0] or {}).keys()) if tbl.pipes else []
    nf = sorted((tbl.nodes[0] or {}).keys()) if tbl.nodes else []
    zf = sorted((tbl.nozzles[0] or {}).keys()) if tbl.nozzles else []
    print(f"\n[표의 칸]  배관 {pf}")
    print(f"           절점 {nf}")
    print(f"           노즐 {zf}")

    # ── 지시서 §3-2 의 속성이 «어디 사는가» — 표인가 kfp 메타인가 둘 다 아닌가.
    #   이것이 §3-4「build_design_tables 직후 한 함수에서 덮는다」가 성립하는지를
    #   가른다. 표에 칸이 없으면 그 함수만으로는 못 덮는다.
    kfp = r1["got"]["kfp"]
    pmeta = sorted((next(iter((kfp.get("pipe_data") or {}).values()), {})
                    or {}).keys())
    nmeta = set()
    for m in (kfp.get("nodes_meta_runtime") or {}).values():
        nmeta |= set((m or {}).keys())
    nmeta = sorted(nmeta)
    WANT = [
        ("배관 · 관종", "type", pf, pmeta),
        ("배관 · 관경", "dia", pf, pmeta),
        ("배관 · 길이", "length", pf, pmeta),
        ("배관 · C", "c", pf, pmeta),
        ("배관 · 거칠기", "roughness_mm", pf, pmeta),
        ("배관 · 등가길이", "equivalent_length", pf, pmeta),
        ("절점 · 표고", "elevation", nf, nmeta),
        ("헤드 · K 값", "k_factor_si", zf, nmeta),
        ("헤드 · 필요압력", "required_pressure_bar", zf, nmeta),
    ]
    print("\n[속성이 사는 곳]  표 = build_design_tables 산출 · 메타 = kfp")
    for label, field, tfields, mfields in WANT:
        in_t = field in tfields
        in_m = field in mfields
        where = ("표+메타" if in_t and in_m else
                 "표" if in_t else "메타" if in_m else "★어느 쪽에도 없음")
        print(f"   {label:<16} {where:<14}"
              f" (표 칸 '{field}' {'O' if in_t else 'X'}"
              f" · kfp 메타 {'O' if in_m else 'X'})")
    print(f"\n   kfp 배관 메타 칸: {pmeta}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

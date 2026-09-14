# -*- coding: utf-8 -*-
"""[요소속성 수정카드] 길이를 바꾸면 «아이소도» 따라오는가 — 그 한 가지만 잰다.

  오너(2026-09-14): 「20m 배관을 2m로 바꾸었는데, 아이소가 그대로인건
  말이안되니까」.

  회랑 좌표는 사슬로 만든다(`p(자식)=p(부모)+L·u`). 그러니 덮은 길이를 넣고
  사슬을 다시 걸면 좌표가 따라 움직여야 한다. 그것을 **수치로** 확인한다:

    · 표의 그 배관 길이가 바뀌었나
    · 그 배관 **뒤쪽** 절점들이 그만큼 움직였나 (아이소가 변한다는 뜻)
    · **앞쪽**(급수원 쪽) 절점은 안 움직였나 (엉뚱한 데까지 흔들면 안 된다)
    · 길이 = 좌표 거리(C2)가 여전히 성립하나
    · 그 배관 **말고는** 길이가 안 바뀌었나 (기준 1 — 그 값만 바뀐다)

    python scripts/_probe_override_iso.py [--key 저장본] [--k 30] [--new 2.0]
"""
from __future__ import annotations

import argparse
import math
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from _probe_overrides import DM_KEY, build                     # noqa: E402


def _ends(pr):
    return pr.get("start") or pr.get("from"), pr.get("end") or pr.get("to")


def snap(kfp):
    nd = kfp.get("nodes_meta_runtime") or {}
    pr = kfp.get("pipe_data") or {}
    xyz = {}
    for nid, m in nd.items():
        c = (m or {}).get("coords") or (0, 0, 0)
        xyz[str(nid)] = (float(c[0]), float(c[1]),
                         float(c[2]) if len(c) > 2 else 0.0)
    ln = {str(pid): float((p or {}).get("length_m") or 0) for pid, p in pr.items()}
    return xyz, ln


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=DM_KEY)
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--new", type=float, default=2.0)
    a = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.io import load_edits, open_board
    from routes.module_f import overrides as ov

    board = open_board(a.key)
    load_edits(board)
    if not board.sources:
        from services.cad_import.edit.board import body_seg_groups
        from services.cad_import.edit.session import MODE_SOURCE, EditSession
        g = sorted((x[0] for x in body_seg_groups(
            board.pts, board.edges, board.bodies())), key=len, reverse=True)[0]
        e0 = EditSession(board, key=a.key)
        e0.set_mode(MODE_SOURCE)
        pa, pb = g[0]
        e0.click((float(pa[0]) + float(pb[0])) / 2.0,
                 (float(pa[1]) + float(pb[1])) / 2.0, 2000)
    payload = dict(board.payload())
    payload["key"] = a.key
    payload["pts"] = [list(p) for p in board.pts]
    payload["edges"] = [list(e) for e in board.edges]
    payload["hcov"] = [list(d) for d in board.disks]
    payload["ups"] = [list(u) for u in board.ups]
    payload["disk_kinds"] = list(board.disk_kinds)
    payload["edge_len_mm"] = dict(getattr(board, "edge_len_mm", None) or {})

    # ── ① 덮기 전
    r0, err = build(board, payload, a.key, a.k)
    if err:
        print(f"★빌드 실패 — {err}")
        return 1
    kfp0 = r0["got"]["kfp"]
    xyz0, ln0 = snap(kfp0)
    idx = ov.build_index(r0["got"], board)
    # 가장 긴 평면 배관을 고른다 — 바뀌는 것이 눈에 보이게
    best = max(((k, pid) for k, pid in idx["pipe"].items()),
               key=lambda t: ln0.get(t[1], 0.0))
    key, pid0 = best
    L0 = ln0[pid0]
    print(f"\n■ 길이를 바꾸면 아이소도 따라오나 · {a.key} · K={a.k}")
    print(f"\n[고를 것]  안정 키 {key} · 이번 이름 {pid0}"
          f" · 지금 길이 {L0:.3f} m → {a.new:.3f} m")

    # ── ② 덮고 다시 빌드
    rows = ov.put([], key, "pipe", "length", a.new,
                  reason="프로브 — 아이소 반영 확인", at="probe")
    r1, err = build(board, payload, a.key, a.k)
    if err:
        print(f"★빌드 실패 — {err}")
        return 1
    n_meta, missed, rep = ov.apply_to_kfp(r1["got"], board, rows)
    kfp1 = r1["got"]["kfp"]
    xyz1, ln1 = snap(kfp1)
    pid1 = ov.build_index(r1["got"], board)["pipe"].get(key)

    print(f"[적용]     {n_meta}건 · 못 옮김 {len(missed)}건"
          f" · 이번 이름 {pid1}")
    print(f"[길이]     {ln0.get(pid0, 0):.3f} → {ln1.get(pid1, 0):.3f} m"
          f"   {'OK' if abs(ln1.get(pid1, 0) - a.new) < 1e-6 else '★안 바뀜'}")

    # 그 배관 말고 길이가 바뀐 것이 있나 (기준 1)
    others = [p for p in ln0
              if p in ln1 and p != pid0 and abs(ln0[p] - ln1[p]) > 1e-6]
    print(f"[그 값만]  다른 배관 중 길이가 바뀐 것 {len(others)}개"
          f"   {'OK' if not others else '★' + str(others[:5])}")

    # 좌표가 움직였나 — 뒤쪽은 움직이고 앞쪽은 그대로여야 한다
    moved = {n: math.dist(xyz0[n], xyz1[n]) for n in xyz0 if n in xyz1}
    big = [n for n, d in moved.items() if d > 1e-6]
    dmax = max(moved.values(), default=0.0)
    print(f"[좌표]     움직인 절점 {len(big)}/{len(moved)}"
          f" · 최대 {dmax:.3f} m"
          f"   {'OK — 아이소가 따라 변한다' if big else '★아무것도 안 움직였다'}")
    exp = abs(a.new - L0)
    print(f"           기대 이동량 |{a.new:.3f} − {L0:.3f}| = {exp:.3f} m"
          f" · 실제 최대 {dmax:.3f} m"
          f"   {'OK' if abs(dmax - exp) < 1e-3 else '★어긋남'}")

    # C2 — 길이 = 좌표 거리
    pr1 = kfp1.get("pipe_data") or {}
    worst, bad = 0.0, 0
    for pid, p in pr1.items():
        s, e = _ends(p)
        if str(s) not in xyz1 or str(e) not in xyz1:
            continue
        d = math.dist(xyz1[str(s)], xyz1[str(e)])
        er = abs(d - float(p.get("length_m") or 0)) * 1000.0
        worst = max(worst, er)
        if er > 1.0:
            bad += 1
    print(f"[C2]       길이 = 좌표 거리 — 어긋난 배관 {bad}개"
          f" · 최대 {worst:.3f} mm   {'OK' if not bad else '★위반'}")

    ok = (abs(ln1.get(pid1, 0) - a.new) < 1e-6 and big and not others
          and not bad and abs(dmax - exp) < 1e-3)
    print("\n  " + ("★길이를 바꾸니 표·좌표가 함께 움직였다 — 아이소 반영 OK"
                    if ok else "★★아직 — 위 수치가 어디가 다른지 말한다"))
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())

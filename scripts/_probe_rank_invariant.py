# -*- coding: utf-8 -*-
"""[최불리규칙복원 §3·§4] 후보가 안 깎였나 · 먼 순서 규칙이 서나.

지시서 `ModuleF_최불리규칙_복원_지시서.md` §3.

    C = 후보 (영역·장 제한을 통과하고 급수원에서 도달하는 헤드 · 자리 대표)
    S = 선정 결과 (|S| = K)

    ①  |S| == K
    ②  min{ far(h) : h ∈ S }  ≥  max{ far(h) : h ∈ C − S }      ← 「먼 순서 그대로」
    ③  S ⊆ C

  ②를 재는 법은 서버와 같다 — 같은 후보로 K+1 개를 뽑으면 늘어난 하나가
  «C−S 의 1등» 이다. 서버가 이미 그 검사를 하고 응답에 실으므로
  (`summary.rank_invariant`), 여기서는 **그 답을 받아서** 확인하고 동시에
  독립으로 한 번 더 잰다 — 검사기 자신이 틀릴 수도 있기 때문이다.

  판 하나를 열어 여러 조합을 이어서 돈다(찍기를 조합마다 다시 하지 않는다).

    python scripts/_probe_rank_invariant.py [--key 저장본] [--k 30]
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from _probe_candidate_drop import _Null, setup, wait      # noqa: E402


def zones_for(disks, n):
    if not n:
        return None
    xs = [float(d[0]) for d in disks]
    ys = [float(d[1]) for d in disks]
    lo, hi = min(xs), max(xs)
    step = (hi - lo) / n
    return [[lo + i * step - 1, min(ys) - 1,
             lo + (i + 1) * step + 1, max(ys) + 1] for i in range(n)]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default="")
    ap.add_argument("--k", type=int, default=0)
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    combos = ([(args.k, 0), (args.k, 2)] if args.k
              else [(30, 0), (30, 2), (12, 0), (12, 2)])
    ctx = _Null() if args.key else isolated_workdir(prefix="rinv_")
    bad = 0
    with ctx, srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        sid = setup(c, argparse.Namespace(key=args.key))

        from routes.module_f.jobs import _sess
        from routes.module_f.remote30 import _worst_k_heads
        sess = _sess(sid)
        b = sess["edit"].board
        est = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (est.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in est["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/mode",
                   json={"sid": sid, "mode": "급수시작위치"})
            c.post("/api/module-f/edit/click",
                   json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                         "y": (seg[1] + seg[3]) / 2, "max_d": 2000})
        src_index = 0 if b.sources else None

        print(f"\n■ 후보·먼순서 검사 · {args.key or '대명동'}"
              f" · 헤드 {len(b.disks)}개")
        for k, nz in combos:
            body = {"sid": sid, "k": k}
            zs = zones_for(b.disks, nz)
            if zs:
                body["zones"] = zs
            jw = c.post("/api/module-f/edit/worst", json=body).get_json()
            if not jw.get("ok"):
                print(f"\n  K={k} 영역{nz or '없음'} · ★최불리 실패 — "
                      f"{str(jw.get('error'))[:140]}")
                bad += 1
                continue
            sm = jw["summary"]
            w = sess["worst"]
            S = [int(h) for h in (w.get("heads") or ())]

            # ── 후보가 «영역 제한만» 받았나 (§0 ① · 수용기준 1)
            only = sess.get("worst_cand")
            cand = (set(range(len(b.disks))) if only is None else
                    {int(i) for i in only})
            full = _worst_k_heads(b.pts, b.edges, b.hnodes, b.sources,
                                  k=max(1, len(cand)), only_heads=cand,
                                  source_index=src_index, head_xy=b.disks)
            C = [int(h) for h in (full.get("heads") or ())]
            far = {int(h): float(v) for h, v in (full.get("dists") or {}).items()}

            i1 = len(S) == k
            i3 = set(S) <= set(C)
            rest = [h for h in C if h not in set(S)]
            mn = min((far[h] for h in S if h in far), default=0.0)
            mx = max((far[h] for h in rest if h in far), default=0.0)
            i2 = mn + 0.1 >= mx
            srv_inv = sm.get("rank_invariant") or {}

            print(f"\n  ── K={k} · 영역 {nz or '없음'}")
            print(f"  [후보]      영역 안 도달 대표 {len(C)}개"
                  f" · 선정에 넘긴 후보 {'도면 전체' if only is None else len(cand)}"
                  f" · 표시 후보 {sm.get('candidates')}")
            print(f"  [불변식]    ① |S|={len(S)} {'OK' if i1 else '★어긋남'}"
                  f" · ② {mn / 1000.0:.3f} ≥ {mx / 1000.0:.3f}"
                  f" {'OK' if i2 else '★어긋남'}"
                  f" · ③ S⊆C {'OK' if i3 else '★어긋남'}")
            print(f"  [서버검사]  {'성립' if srv_inv.get('ok') else '★' + str(srv_inv.get('violations'))}"
                  + (f" (min {srv_inv.get('min_in')} ≥ max {srv_inv.get('max_out')})"
                     if srv_inv.get("min_in") is not None else ""))
            na = sm.get("not_attachable") or {}
            print(f"  [안 붙는 것] 후보 중 {na.get('n', 0)}개 — "
                  f"{na.get('by_why') or '(없음)'}  ※빼지 않았다")
            if not (i1 and i2 and i3 and srv_inv.get("ok")):
                bad += 1
                for h in rest[:3]:
                    print(f"      안 뽑힌 상위 disk {h}"
                          f" {far.get(h, 0) / 1000.0:.2f} m")
        print("\n  " + ("★§3 불변식이 모든 조합에서 선다 — 규칙 복원됨"
                        if not bad else f"★★{bad}개 조합에서 안 선다"))
    return 0 if not bad else 2


if __name__ == "__main__":
    raise SystemExit(main())

# -*- coding: utf-8 -*-
"""[요소속성 수정카드] 라우트로 고치면 표·.sdf·아이소가 따라오나 — 끝에서 끝까지.

  화면이 밟는 길을 그대로 밟는다(HTTP):

      표 확정 → 카드가 값을 본다 → POST /design/override → 표 재확정
      → 표 값이 바뀌었나 · 좌표가 움직였나 · 파일에 남았나 · 그 값만 바뀌었나

    python scripts/_probe_override_e2e.py [--key 저장본] [--k 30]
"""
from __future__ import annotations

import argparse
import json
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

from _probe_candidate_drop import _Null, setup, wait            # noqa: E402


def tbl_snapshot(sess):
    d = sess.get("design") or {}
    tbl = d.get("tables")
    pipes = {str(r.get("label")): (float(r.get("length") or 0),
                                   r.get("dia"))
             for r in (getattr(tbl, "pipes", None) or ())}
    nodes = {str(n.get("label")): (float(n.get("x") or 0),
                                   float(n.get("y") or 0),
                                   float(n.get("elevation") or 0))
             for n in (getattr(tbl, "nodes", None) or ())}
    return pipes, nodes


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default="")
    ap.add_argument("--k", type=int, default=30)
    a = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    ctx = _Null() if a.key else isolated_workdir(prefix="ove2e_")
    with ctx, srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        sid = setup(c, argparse.Namespace(key=a.key))
        from routes.module_f.jobs import _sess
        from routes.module_f import overrides as ov
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
        jw = c.post("/api/module-f/edit/worst",
                    json={"sid": sid, "k": a.k}).get_json()
        if not jw.get("ok"):
            print(f"★최불리 실패 — {str(jw)[:200]}")
            return 1

        def rebuild():
            c.post("/api/module-f/design/build", json={"sid": sid, "k": a.k})
            return wait(c, sid).get("state") == "done"

        if not rebuild():
            print("★표 확정 실패")
            return 1
        p0, n0 = tbl_snapshot(sess)
        print(f"\n■ 요소 수정 끝에서 끝까지 · {a.key or '대명동'} · K={a.k}")
        print(f"\n[표]      배관 {len(p0)} · 절점 {len(n0)}")

        # ── 가장 긴 평면 배관을 고른다
        idx = ov.build_index(sess["design"]["got"], b)
        best = None
        for key, pid in idx["pipe"].items():
            row = p0.get(str(pid))
            if row and (best is None or row[0] > best[2]):
                best = (key, pid, row[0])
        if best is None:
            print("★고칠 배관을 못 찾음")
            return 1
        key, pid, L0 = best
        new_len = round(L0 / 5.0, 3)
        print(f"[고른 것] 안정 키 {key} · 이번 이름 {pid}"
              f" · 길이 {L0:.3f} → {new_len:.3f} m")

        # ── 사유 없이 보내면 거절해야 한다 (기준 8 · 길이는 사유 필수)
        r = c.post("/api/module-f/design/override",
                   json={"sid": sid, "kind": "pipe", "key": list(key),
                         "field": "length", "new": new_len}).get_json()
        print(f"[사유없이] {'★통과됨(문제)' if r.get('ok') else 'ok — 거절'}"
              f" · {str(r.get('message'))[:70]}")

        # ── 범위 밖 값도 거절
        r = c.post("/api/module-f/design/override",
                   json={"sid": sid, "kind": "pipe", "key": list(key),
                         "field": "length", "new": -3,
                         "reason": "시험"}).get_json()
        print(f"[음수]     {'★통과됨(문제)' if r.get('ok') else 'ok — 거절'}"
              f" · {str(r.get('message'))[:70]}")

        # ── 제대로 저장
        r = c.post("/api/module-f/design/override",
                   json={"sid": sid, "kind": "pipe", "key": list(key),
                         "field": "length", "new": new_len,
                         "reason": "현장 실측"}).get_json()
        print(f"[저장]     {'ok' if r.get('ok') else '★실패'}"
              f" · 파일 {r.get('file')} · 수정 {r.get('count')}건")

        if not rebuild():
            print("★재확정 실패")
            return 1
        p1, n1 = tbl_snapshot(sess)
        pid1 = ov.build_index(sess["design"]["got"], b)["pipe"].get(key)
        g1 = p1.get(str(pid1))
        print(f"[표 길이]  {L0:.3f} → {g1[0]:.3f} m"
              f"   {'OK' if abs(g1[0] - new_len) < 1e-3 else '★안 바뀜'}")
        others = [k for k in p0
                  if k in p1 and k != str(pid) and abs(p0[k][0] - p1[k][0]) > 1e-6]
        print(f"[그 값만]  다른 배관 길이가 바뀐 것 {len(others)}개"
              f"   {'OK' if not others else '★' + str(others[:4])}")
        moved = [k for k in n0 if k in n1
                 and math.dist(n0[k][:2], n1[k][:2]) > 1e-6]
        print(f"[좌표]     움직인 절점 {len(moved)}/{len(n0)}"
              f"   {'OK — 아이소가 따라 변한다' if moved else '★안 움직임'}")

        # ── 파일에 남았나 (D2)
        fp = ov.path_for(sess.get("key") or "design")
        on_disk = []
        if os.path.exists(fp):
            on_disk = (json.load(open(fp, encoding="utf-8")) or {}).get("items") or []
        print(f"[파일]     {os.path.basename(fp)} · {len(on_disk)}건"
              f"   {'OK' if on_disk else '★안 남음'}")
        if on_disk:
            it = on_disk[0]
            print(f"           원값 {it.get('old')} → 새값 {it.get('new')}"
                  f" · 사유 «{it.get('reason')}» · {it.get('at')}"
                  f"   {'OK — 원값·사유·시각이 남는다' if it.get('old') is not None else '★원값 없음'}")

        # ── 지우면 원값으로 (기준 4)
        c.post("/api/module-f/design/override",
               json={"sid": sid, "kind": "pipe", "key": list(key),
                     "field": "length", "remove": True})
        if not rebuild():
            print("★재확정 실패")
            return 1
        p2, _n2 = tbl_snapshot(sess)
        pid2 = ov.build_index(sess["design"]["got"], b)["pipe"].get(key)
        back = p2.get(str(pid2))
        print(f"[지우기]   {back[0]:.3f} m"
              f"   {'OK — 원값으로 돌아왔다' if abs(back[0] - L0) < 1e-3 else '★안 돌아옴'}")

        ok = (abs(g1[0] - new_len) < 1e-3 and not others and moved
              and on_disk and abs(back[0] - L0) < 1e-3)
        print("\n  " + ("★끝에서 끝까지 선다" if ok else "★★아직 — 위 수치를 본다"))
        return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())

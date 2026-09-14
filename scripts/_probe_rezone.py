# -*- coding: utf-8 -*-
"""최불리 → 해제 → 영역 다시 → 최불리 가 되는가 (2026-09-14 사용자 지적).

  사용자: 「최초 최불리 선정 후에 해제하고 영역 다시 지정 후에 최불리 선정을
  누르니까 작동을 안 한다」. 화면 말고 **라우트 순서를 그대로 밟아** 재현한다.

    python scripts/_probe_rezone.py [--key 저장본] [--k 30]
"""
from __future__ import annotations

import argparse
import io as _io
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from _probe_candidate_drop import _Null, setup, wait          # noqa: E402


def zones_half(disks, lo_frac, hi_frac):
    xs = [float(d[0]) for d in disks]
    ys = [float(d[1]) for d in disks]
    lo, hi = min(xs), max(xs)
    w = hi - lo
    return [[lo + w * lo_frac - 1, min(ys) - 1,
             lo + w * hi_frac + 1, max(ys) + 1]]


def step(c, sid, label, body):
    t0 = time.perf_counter()
    j = c.post("/api/module-f/edit/worst", json=body).get_json() or {}
    dt = time.perf_counter() - t0
    ok = bool(j.get("ok"))
    s = j.get("summary") or {}
    print(f"  [{label}] {'ok' if ok else '★실패'} · {dt:.1f}s"
          + (f" · K={s.get('k')} · 후보 {s.get('candidates')}"
             f" · 영역 {s.get('zones')}" if ok else
             f" · {str(j.get('message') or j.get('error'))[:160]}"))
    return j


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

    ctx = _Null() if a.key else isolated_workdir(prefix="rez_")
    with ctx, srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        sid = setup(c, argparse.Namespace(key=a.key))
        from routes.module_f.jobs import _sess
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

        zA = zones_half(b.disks, 0.0, 0.55)
        zB = zones_half(b.disks, 0.45, 1.0)
        print(f"\n■ 영역 바꿔 가며 최불리 · {a.key or '대명동'} · K={a.k}")

        step(c, sid, "① 영역A 로 최불리", {"sid": sid, "k": a.k, "zones": zA})
        print(f"      세션: worst={bool(sess.get('worst'))}"
              f" · zones={len(sess.get('worst_zones') or ())}"
              f" · cand={len(sess.get('worst_cand') or ()) if sess.get('worst_cand') is not None else '전체'}")

        cl = c.post("/api/module-f/edit/worst-clear",
                    json={"sid": sid}).get_json() or {}
        print(f"  [② 해제] {'ok' if cl.get('ok') else '★실패'}")
        print(f"      세션: worst={bool(sess.get('worst'))}"
              f" · zones={len(sess.get('worst_zones') or ())}"
              f" · cand={len(sess.get('worst_cand') or ()) if sess.get('worst_cand') is not None else '전체'}"
              f" · args={bool(sess.get('worst_args'))}")

        step(c, sid, "③ 영역B 로 최불리", {"sid": sid, "k": a.k, "zones": zB})
        step(c, sid, "④ 영역 없이 최불리", {"sid": sid, "k": a.k})
        step(c, sid, "⑤ 영역A 로 다시", {"sid": sid, "k": a.k, "zones": zA})
        return 0


if __name__ == "__main__":
    raise SystemExit(main())

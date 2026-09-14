# -*- coding: utf-8 -*-
"""[최불리규칙복원 §1] 후보 좁히기가 «상위 K» 를 떨어뜨렸는가 — 고치기 전에 잰다.

지시서 `ModuleF_최불리규칙_복원_지시서.md` §1.

  규칙은 한 줄이다 — 「정의된 헤드 전부를 후보로 두고 유하거리가 긴 순서
  그대로 K 개」. `/edit/worst` 가 고르기 전에 `only = before & keep` 으로
  후보를 깎으면 `worst_k_heads` 는 그 헤드를 **순위 계산에서 아예 뺀다**::

      for hi, nodes in enumerate(hnodes):
          if only_heads is not None and hi not in only_heads:
              continue

  그래서 빠진 헤드가 1등이었다면 2등이 1등 자리에 온다. 그 일이 실제로
  일어나는지를 **고치기 전에** 숫자로 확정한다. 이 수가 0 이 아니면 §1 확정.

    python scripts/_probe_candidate_drop.py [--k 30] [--zones N] [--key 저장본]
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
           str(ROOT / "tests")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DXF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=200000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


class _Null:
    def __enter__(self): return None
    def __exit__(self, *a): return False


def setup(c, args):
    """열기 → 찍기 → 손질 (또는 저장본 이어 열기). sid 를 돌려준다."""
    if args.key:
        j0 = c.post("/api/module-f/reopen", json={"key": args.key}).get_json()
        if not (j0 or {}).get("ok"):
            raise SystemExit(f"★이어서 열기 실패 — {str(j0)[:200]}")
        sid = j0["sid"]
        if wait(c, sid).get("state") != "done":
            raise SystemExit("★열기 잡 실패")
        return sid
    with open(DXF, "rb") as f:
        raw = f.read()
    sid = c.post("/api/module-f/open", data={
        "dxf_file": (_io.BytesIO(raw), os.path.basename(str(DXF)))},
        content_type="multipart/form-data").get_json()["sid"]
    wait(c, sid)
    c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
    c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
    c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "complete"})
    c.post("/api/module-f/pick/mode",
           json={"sid": sid, "action": "slot", "slot": "상향"})
    c.post("/api/module-f/pick/suggest", json={"sid": sid})
    wait(c, sid)
    cands = ((c.get(f"/api/module-f/convert/result?sid={sid}")
              .get_json()["result"] or {}).get("candidates") or [])
    for c_ in cands:
        d = c.post("/api/module-f/pick/click",
                   json={"sid": sid, "x": c_["x"], "y": c_["y"],
                         "max_d": 300}).get_json()
        if (d.get("report") or {}).get("동작") == "취소":
            c.post("/api/module-f/pick/click",
                   json={"sid": sid, "x": c_["x"], "y": c_["y"],
                         "max_d": 300})
    c.post("/api/module-f/pick/commit", json={"sid": sid})
    wait(c, sid)
    return sid


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--zones", type=int, default=0)
    ap.add_argument("--key", default="")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    ctx = _Null() if args.key else isolated_workdir(prefix="cdrop_")
    with ctx, srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        sid = setup(c, args)

        from routes.module_f.jobs import _sess
        from routes.module_f.remote30 import _worst_k_heads
        from routes.module_f.attach import wet_heads
        sess = _sess(sid)
        es = sess["edit"]
        b = es.board
        est = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (est.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in est["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/mode",
                   json={"sid": sid, "mode": "급수시작위치"})
            c.post("/api/module-f/edit/click",
                   json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                         "y": (seg[1] + seg[3]) / 2, "max_d": 2000})

        # ── 후보: 영역 제한만 건다 (§0 — 그 밖의 이유로는 줄이지 않는다)
        before = set(range(len(b.disks)))
        if args.zones:
            xs = [float(d[0]) for d in b.disks]
            ys = [float(d[1]) for d in b.disks]
            lo, hi = min(xs), max(xs)
            step = (hi - lo) / args.zones
            rects = [(lo + i * step - 1, min(ys) - 1,
                      lo + (i + 1) * step + 1, max(ys) + 1)
                     for i in range(args.zones)]
            before = {i for i in before
                      if any(x0 <= float(b.disks[i][0]) <= x1
                             and y0 <= float(b.disks[i][1]) <= y1
                             for x0, y0, x1, y1 in rects)}

        # ── 좁히기가 뺐을 헤드 (지금 코드가 하는 것과 같은 계산)
        probe = wet_heads(sess, es)
        wet = {int(i) for i in (probe.get("wet") or ())}
        shared = {int(i) for i in (probe.get("shared") or ())}
        keep = wet - shared
        dropped = before - keep
        reason = probe.get("reason") or {}

        def why(i):
            return str(reason.get(i) or reason.get(str(i))
                       or ("shared" if i in shared else "unknown"))

        def count_by(vals):
            out: dict = {}
            for v in vals:
                out[v] = out.get(v, 0) + 1
            return dict(sorted(out.items(), key=lambda t: -t[1]))

        src_index = 0 if b.sources else None

        def rank(only, k):
            return _worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=k,
                                  only_heads=only, source_index=src_index,
                                  head_xy=b.disks)

        # ── ㉠ 규칙대로: 후보 전부를 순위에 넣는다
        full = rank(before, len(before) or 1)
        order = list(full["heads"])            # 유하거리 긴 순서 (대표만)
        top = order[:args.k]
        far = dict(full.get("dists") or {})

        # ── ㉡ 지금 코드: 좁힌 뒤 고른다
        narrow = rank(before & keep, args.k)
        picked = list(narrow["heads"])

        lost = [h for h in top if h in dropped]
        print(f"\n■ 후보 좁히기가 상위 K 를 떨어뜨렸나 · K={args.k}"
              f" · 영역 {args.zones or '없음'} · {args.key or '대명동'}")
        print(f"\n[후보]      영역 안 헤드 {len(before)}개"
              f" → 좁힌 뒤 {len(before & keep)}개"
              f" · 도달 대표 {full['reachable']}개")
        print(f"[원인확인]  후보 좁히기로 빠졌던 헤드 {len(dropped)}개"
              f" · 그중 상위 K 안 {len(lost)}개"
              + ("   ★★§1 확정 — 규칙이 깨졌다" if lost else
                 "   (0 — 이 판에서는 순위에 영향 없음)"))
        print(f"[사유]      " + (" · ".join(f"{k2} {v}" for k2, v in
                                            count_by(why(i) for i in dropped)
                                            .items()) or "(없음)"))
        if lost:
            print(f"[떨어진 상위 K]  자리 · 유하거리 · 사유"
                  f"  (1등부터 {min(len(lost), 12)}개)")
            for h in lost[:12]:
                r = order.index(h) + 1
                print(f"    {r:>4}등  disk {h:<5}"
                      f" ({float(b.disks[h][0]):.0f}, {float(b.disks[h][1]):.0f})"
                      f" · {far.get(h, 0) / 1000.0:>7.2f} m · {why(h)}")
            took = [h for h in picked if h not in top]
            print(f"[대신 들어온 것] {len(took)}개"
                  f"  (규칙대로면 K 밖이어야 한다)")
            for h in took[:12]:
                r = (order.index(h) + 1) if h in order else -1
                print(f"    {r:>4}등  disk {h:<5}"
                      f" ({float(b.disks[h][0]):.0f}, {float(b.disks[h][1]):.0f})"
                      f" · {far.get(h, 0) / 1000.0:>7.2f} m")
        print(f"[1등]       규칙대로 {full.get('far_m')} m"
              f" · 좁힌 뒤 {narrow.get('far_m')} m"
              + ("   ★최원 유하거리가 바뀌었다"
                 if full.get("far_m") != narrow.get("far_m") else ""))
        return 0 if not lost else 2


if __name__ == "__main__":
    raise SystemExit(main())

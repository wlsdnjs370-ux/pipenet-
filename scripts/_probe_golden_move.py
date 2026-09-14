# -*- coding: utf-8 -*-
"""골든의 `daemyeong.worst` 가 왜 움직였나 — 옛 규칙과 새 규칙을 나란히 잰다.

  `tests/module_f_complete_golden.json` 의 그 칸은 스스로 「움직였다면 다른
  이야기이므로 멈추고 물어라」고 적어 둔 자리다. 그래서 덮기 전에, **같은
  판에서** 두 규칙을 각각 돌려 무엇이 바뀌었는지 헤드 단위로 낸다:

      옛 규칙  후보 = 영역 안 전부          ← 표가 못 붙이는 헤드가 섞인다
      새 규칙  후보 = 표에 노즐로 오는 헤드  ← `손질정본 §2-1`

  판정 규칙은 건드리지 않는다. **어느 후보 집합을 넘기느냐**만 다르다.

    python scripts/_probe_golden_move.py [--k 10]
"""
from __future__ import annotations

import argparse
import io as _io
import math
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


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=10)
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with isolated_workdir(prefix="gmove_"), srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(DXF, "rb") as f:
            raw = f.read()
        sid = c.post("/api/module-f/open", data={
            "dxf_file": (_io.BytesIO(raw), os.path.basename(str(DXF)))},
            content_type="multipart/form-data").get_json()["sid"]
        wait(c, sid)
        c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
        c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
        c.post("/api/module-f/pick/mode",
               json={"sid": sid, "action": "complete"})
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

        def pick(only):
            return _worst_k_heads(b.pts, b.edges, b.hnodes, b.sources,
                                  k=args.k, only_heads=only,
                                  source_index=0, head_xy=b.disks)

        probe = wet_heads(sess, es)
        wet = {int(i) for i in (probe.get("wet") or ())}
        shared = {int(i) for i in (probe.get("shared") or ())}
        keep = wet - shared
        reason = probe.get("reason") or {}

        old = pick(None)
        new = pick(keep)
        so, sn = set(old["heads"]), set(new["heads"])

        def why(i):
            return str(reason.get(i) or reason.get(str(i))
                       or ("shared" if i in shared else "unknown"))

        print(f"\n■ 골든이 움직인 이유 · 대명동 · K={args.k}")
        print(f"\n[후보]     도면 헤드 {len(b.disks)}"
              f" → 표에 노즐로 오는 헤드 {len(keep)}"
              f" (못 붙는 {len(b.disks) - len(keep)})")
        for lbl, w2 in (("옛 규칙", old), ("새 규칙", new)):
            print(f"[{lbl}]   총연장 {w2.get('total_m', 0):.2f} m"
                  f" · 최원 {w2.get('far_m', 0):.2f} m"
                  f" · 폭 {w2.get('span_m', 0):.2f} m"
                  f" · 최대담당 {w2.get('max_load', 0)}")
        print(f"[선정 차]  같은 헤드 {len(so & sn)}"
              f" · 옛것에만 {len(so - sn)} · 새것에만 {len(sn - so)}")
        for i in sorted(so - sn):
            x, y = float(b.disks[i][0]), float(b.disks[i][1])
            print(f"    옛것에만 disk {i} ({x:.0f}, {y:.0f})"
                  f" · 표에 오나? {'예' if i in keep else '아니오 — ' + why(i)}")
        for i in sorted(sn - so):
            x, y = float(b.disks[i][0]), float(b.disks[i][1])
            near = min(((math.dist((x, y), (float(b.disks[j][0]),
                                            float(b.disks[j][1]))), j)
                        for j in (so - sn)), default=(0.0, -1))
            print(f"    새것에만 disk {i} ({x:.0f}, {y:.0f})"
                  f" · 옛것에만 있던 헤드까지 {near[0]:.0f} mm")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())

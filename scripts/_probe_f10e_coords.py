# -*- coding: utf-8 -*-
"""[F-10e] 아이소 밑그림을 깔려면 «어느 좌표계에서 어디로» 가는지 알아야 한다.

지시서는 밑그림을 corridor 와 **같은 변환**으로 만들라고 못 박았다. 그런데
엔진의 `normalize_node_coords` 는 제 노드 집합의 bbox 로 배율을 정하므로,
board 를 따로 넣으면 배율이 갈린다 — 겹치지 않는다.

그래서 먼저 사실을 잰다:

  ① 설계 표의 노드 좌표가 board 세계 mm 인가 (변환 «전» 값이 남아 있나)
  ② 그 노드가 board 절점과 실제로 같은 자리인가
  ③ 변환 «후»(display_tables) 좌표와 짝을 지을 수 있나
  ④ 노드 표고 분포 — 밑그림에 쓸 «대표 표고(최빈값)» 가 하나로 잡히나

    python scripts/_probe_f10e_coords.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=3000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEF
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True, "heads": {"conf_min": 0.9}})
        wait(c, sid)
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        # 헤드 가까운 배관 끝점에 원클릭
        pipe = []
        for g in st["body_groups"]:
            s2 = g["segs"]
            for i in range(0, len(s2) - 3, 4):
                pipe.append((float(s2[i]), float(s2[i + 1])))
        hx, hy = float(st["heads"][0][0]), float(st["heads"][0][1])
        px, py = min(pipe, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": px, "y": py})
        j = wait(c, sid)
        print(f"원클릭 {j['state']}  {j.get('error') or ''}")
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
        j = wait(c, sid)
        print(f"표 확정 {j['state']}  {j.get('error') or ''}")
        if j["state"] != "done":
            return 2

        from routes.module_f.jobs import _SESSIONS
        sess = _SESSIONS[sid]
        tbl = sess["design"]["tables"]
        board = sess["edit"].board

        print("\n■ ① 설계 표 노드의 «변환 전» 좌표")
        xs = [float(n.get("x", 0) or 0) for n in tbl.nodes]
        ys = [float(n.get("y", 0) or 0) for n in tbl.nodes]
        print(f"    노드 {len(tbl.nodes)} · x {min(xs):,.0f}~{max(xs):,.0f} · "
              f"y {min(ys):,.0f}~{max(ys):,.0f}")
        bxs = [float(p[0]) for p in board.pts]
        bys = [float(p[1]) for p in board.pts]
        print(f"    board 절점 {len(board.pts)} · x {min(bxs):,.0f}~{max(bxs):,.0f} · "
              f"y {min(bys):,.0f}~{max(bys):,.0f}")

        print("\n■ ② 설계 노드가 board 절점과 같은 자리인가")
        bset = {(round(float(p[0]), 3), round(float(p[1]), 3)) for p in board.pts}
        hit = sum(1 for n in tbl.nodes
                  if (round(float(n.get("x", 0) or 0), 3),
                      round(float(n.get("y", 0) or 0), 3)) in bset)
        print(f"    board 절점과 정확히 일치 {hit} / {len(tbl.nodes)}")

        print("\n■ ③ 변환 «후» 와 짝짓기")
        pv = c.get(f"/api/module-f/design/preview?sid={sid}&iso=1").get_json()
        vn = {str(n["label"]): n for n in pv["view"]["nodes"]}
        print(f"    preview 노드 {len(vn)} · 라벨로 짝지어짐 "
              f"{sum(1 for n in tbl.nodes if str(n.get('label')) in vn)}")
        vxs = [float(n["x"]) for n in pv["view"]["nodes"]]
        vys = [float(n["y"]) for n in pv["view"]["nodes"]]
        print(f"    아이소 x {min(vxs):,.1f}~{max(vxs):,.1f} · "
              f"y {min(vys):,.1f}~{max(vys):,.1f}")

        print("\n■ ④ 표고 분포 — 밑그림에 쓸 대표 표고")
        ev = Counter(round(float(n.get("elevation", 0) or 0), 3)
                     for n in tbl.nodes)
        for e, k in ev.most_common(5):
            print(f"    {e:>10,.3f} m  ·  {k}개")
        top, cnt = ev.most_common(1)[0]
        print(f"    최빈값 {top} ({cnt}/{len(tbl.nodes)} = "
              f"{cnt / len(tbl.nodes) * 100:.0f}%)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

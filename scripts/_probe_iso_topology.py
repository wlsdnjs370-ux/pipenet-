# -*- coding: utf-8 -*-
"""[아이소 위상] 화면이 약속한 「갈림·끝·고리 수는 그대로」가 지켜지나.

■ 화면이 사람에게 한 약속 (templates/module_f.html · 수리계산 패널)

    「아이소는 **같은 망**을 설계 좌표로 옮겨 그립니다. 다르게 보이는 것은
     둘뿐입니다 — 헤드를 **세워** 그리므로 평면에 없는 짧은 토막이 헤드마다
     생기고, 일직선 중간 절점은 병합되어 절점 수가 줍니다
     (**갈림·끝·고리 수는 그대로**).」

  약속에 수가 셋 박혀 있으니 그대로 센다. 「달라 보인다」는 인상을 인상으로
  두지 않고, 어느 수가 어긋나는지로 바꾼다.

■ 무엇을 무엇과 맞대나

    ㉠ 손질망 중 «설계면적이 쓰는 부분»  = 찍기에서 수동 지정한 배관·헤드 그대로
       (corridor 가 덮는 board 간선을 그 주변까지 넓혀 잡는다)
    ㉡ 표(= 아이소가 그리는 망)          = 제한 전개 산출

  위상 수 세 가지를 양쪽에서 센다:
      갈림 = 차수 >= 3 절점 · 끝 = 차수 1 절점 · 고리 = E - V + C

    python scripts/_probe_iso_topology.py [--k 30] [--zones]
"""
from __future__ import annotations

import argparse
import io as _io
import math
import os
import sys
import time
from collections import defaultdict, deque
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


def topo(edges) -> dict:
    """간선 목록 → 갈림·끝·고리·성분. 간선은 (a, b) 이름쌍."""
    deg = defaultdict(int)
    adj = defaultdict(set)
    ue = set()
    for a, b in edges:
        if a == b:
            continue
        ue.add((a, b) if a <= b else (b, a))
    for a, b in ue:
        deg[a] += 1
        deg[b] += 1
        adj[a].add(b)
        adj[b].add(a)
    seen, comp = set(), 0
    for n in list(adj):
        if n in seen:
            continue
        comp += 1
        q = deque([n])
        seen.add(n)
        while q:
            u = q.popleft()
            for v in adj[u]:
                if v not in seen:
                    seen.add(v)
                    q.append(v)
    V, E = len(deg), len(ue)
    return {"절점": V, "간선": E,
            "갈림": sum(1 for d in deg.values() if d >= 3),
            "끝": sum(1 for d in deg.values() if d == 1),
            "성분": comp, "고리": E - V + comp}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--zones", action="store_true")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with isolated_workdir(prefix="isotopo_"), srv.app.test_client() as c:
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

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        b = sess["edit"].board
        est = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        seg = est["body_groups"][0]["segs"]
        c.post("/api/module-f/edit/mode",
               json={"sid": sid, "mode": "급수시작위치"})
        c.post("/api/module-f/edit/click",
               json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                     "y": (seg[1] + seg[3]) / 2, "max_d": 2000})

        body = {"sid": sid, "k": args.k}
        if args.zones:
            ds = [(float(d[0]), float(d[1])) for d in b.disks]
            xs = [p[0] for p in ds]; ys = [p[1] for p in ds]
            mid = (min(xs) + max(xs)) / 2
            body["zones"] = [[min(xs) - 1, min(ys) - 1, mid, max(ys) + 1],
                             [mid, min(ys) - 1, max(xs) + 1, max(ys) + 1]]
        jw = c.post("/api/module-f/edit/worst", json=body).get_json()
        if not jw.get("ok"):
            print(f"★최불리 실패 — {str(jw)[:200]}")
            return 1
        w = sess["worst"]
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        if wait(c, sid).get("state") != "done":
            print("★표 확정 실패")
            return 1
        got, tbl = sess["design"]["got"], sess["design"]["tables"]

        print(f"\n■ 아이소 위상 계약 검사 · 대명동 · K={args.k}"
              f" · 영역 {'2곳' if args.zones else '없음'}")

        # ㉠ 손질망에서 «설계면적이 쓰는 부분» — corridor 그대로
        cor = [(int(a), int(c2)) for (a, c2) in (w.get("loads") or {})]
        t_edit = topo(cor)

        # ㉡ 표(아이소) — 세로 구간(헤드 스텁)은 약속이 «생긴다» 고 한 것이라
        #    위상 비교에서는 평면 길이 0 인 것을 접는다.
        origin = got.get("origin_mm") or (0.0, 0.0)
        ox, oy = float(origin[0]) - 1000.0, float(origin[1]) - 1000.0
        at = {str(n.get("label")): (float(n.get("x") or 0) + ox,
                                    float(n.get("y") or 0) + oy)
              for n in tbl.nodes}
        # ★세로 구간을 **버리면 안 된다.** 평면에서 길이가 0 이라고 지우면
        #   헤드가 망에서 떨어져 나가 성분이 1 → 31 로 부풀고 끝이 +42 로 는다
        #   (실측으로 한 번 그렇게 잘못 셌다). 평면에서 «점 하나» 라는 말은
        #   **두 절점을 합친다** 는 뜻이다 — 합치고 센다.
        parent: dict = {}

        def find(x):
            parent.setdefault(x, x)
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x

        def union(x, y):
            rx, ry = find(x), find(y)
            if rx != ry:
                parent[rx] = ry

        vert = 0
        for pr in tbl.pipes:
            a, z = str(pr.get("in")), str(pr.get("out"))
            pa, pz = at.get(a), at.get(z)
            if pa and pz and math.dist(pa, pz) <= 1.0:
                vert += 1
                union(a, z)          # 평면에서는 점 하나 — 합친다
        flat = []
        for pr in tbl.pipes:
            a, z = str(pr.get("in")), str(pr.get("out"))
            ra, rz = find(a), find(z)
            if ra != rz:
                flat.append((ra, rz))
        t_tbl = topo(flat)

        print(f"\n  {'항목':<8}{'㉠ 손질(평면)':>16}{'㉡ 표(아이소)':>16}{'차이':>10}")
        print("  " + "─" * 50)
        for k in ("절점", "간선", "갈림", "끝", "성분", "고리"):
            d = t_tbl[k] - t_edit[k]
            mark = "" if d == 0 else "  ★"
            print(f"  {k:<9}{t_edit[k]:>14}{t_tbl[k]:>16}{d:>+10}{mark}")
        print(f"  {'헤드':<9}{len(w.get('heads') or ()):>14}"
              f"{len(tbl.nozzles):>16}"
              f"{len(tbl.nozzles) - len(w.get('heads') or ()):>+10}")
        print(f"  (세로 구간 {vert}개는 평면에서 점 하나 — 약속대로 «생기는» 것)")

        bad = [k for k in ("갈림", "끝", "고리") if t_tbl[k] != t_edit[k]]
        print()
        if bad:
            print(f"  ★★약속이 깨졌다 — {' · '.join(bad)} 가 다르다.")
            print("     화면은 「갈림·끝·고리 수는 그대로」라고 적어 두었다.")
        else:
            print("  ★약속대로다 — 갈림·끝·고리가 같다(절점 수만 병합으로 준다).")
        return 2 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())

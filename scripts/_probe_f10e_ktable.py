# -*- coding: utf-8 -*-
"""[F-10e] K 를 키운 표가 corridor 와 «같은 좌표 공간» 인가.

앞선 두 실측으로 밝혀진 것:
  · 설계 표 좌표는 board 세계 mm 가 아니다(일치 0/178).
  · board → 설계 표 사이에 전역 아핀이 없다(최대 잔차 9.3%). 엔진이 노드를
    개별로 옮긴다(「직선 위치 복원」).

그러면 밑그림을 board 에서 가져올 수 없다. 남은 길은 «같은 엔진이 만든 더 넓은
표» 를 밑그림으로 쓰는 것뿐이다. 그러려면 K 를 키운 표에서 **공유 노드가 같은
자리** 여야 한다 — 아니면 그것도 어긋난다.

대응은 `node_ref`(설계 노드 ↔ board 절점)로 짓는다. 라벨은 빌드마다 다시
매겨질 수 있으므로 라벨로 짝지으면 안 된다.

    python scripts/_probe_f10e_ktable.py [도면.dxf]
"""
from __future__ import annotations

import os
import statistics
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=6000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def snapshot(sess):
    """{board 절점 index: (x, y)} — 설계 표의 «변환 전» 좌표."""
    tbl = sess["design"]["tables"]
    got = sess["design"]["got"]
    nref = got.get("node_ref") or {}
    if not snapshot._shown:
        snapshot._shown = True
        items = list(nref.items())[:6]
        print(f"    node_ref 표본: {items}")
        print(f"    origin_mm = {got.get('origin_mm')}")
        print(f"    설계 라벨 표본: {[str(n.get('label')) for n in tbl.nodes[:6]]}")
    at = {str(n.get("label")): n for n in tbl.nodes}
    out = {}
    for lab, bi in nref.items():
        # node_ref 는 «N1» 꼴이고 표 라벨은 «1» 이다 — 접두사만 다르다.
        key = str(lab)
        n = at.get(key) or at.get(key[1:] if key[:1] in "Nn" else key)
        if n is None:
            continue
        try:
            k = int(bi)
        except (TypeError, ValueError):
            continue
        out[k] = (float(n.get("x", 0) or 0), float(n.get("y", 0) or 0))
    return out, len(tbl.nodes), dict(tbl.meta).get("앵커 노드")


snapshot._shown = False


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
        pipe = []
        for g in st["body_groups"]:
            s2 = g["segs"]
            for i in range(0, len(s2) - 3, 4):
                pipe.append((float(s2[i]), float(s2[i + 1])))
        hx, hy = float(st["heads"][0][0]), float(st["heads"][0][1])
        px, py = min(pipe, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": px, "y": py})
        wait(c, sid)

        from routes.module_f.jobs import _SESSIONS
        sess = _SESSIONS[sid]

        snaps = {}
        for k in (30, 200):
            c.post("/api/module-f/design/build", json={"sid": sid, "k": k})
            j = wait(c, sid)
            if j["state"] != "done":
                print(f"K={k} 표 확정 실패 — {j.get('error')}")
                return 2
            snaps[k] = snapshot(sess)
            print(f"  K={k:<4} 설계 노드 {snaps[k][1]:>5} · "
                  f"board 대응 {len(snaps[k][0]):>5} · 앵커 {snaps[k][2]}")

        a, b = snaps[30][0], snaps[200][0]
        both = sorted(set(a) & set(b))
        print(f"\n■ 두 표가 공유하는 board 절점 {len(both)}개")
        if not both:
            print("  ★공유가 없다 — 이 길도 막혔다")
            return 3
        d = [((a[k][0] - b[k][0]) ** 2 + (a[k][1] - b[k][1]) ** 2) ** 0.5
             for k in both]
        d.sort()
        print(f"  같은 절점의 좌표 차 — 중앙값 {statistics.median(d):,.3f} · "
              f"p90 {d[int(len(d) * .9)]:,.3f} · 최대 {d[-1]:,.3f}")
        same = sum(1 for v in d if v < 1e-6)
        print(f"  정확히 같은 자리 {same} / {len(d)} "
              f"({same / len(d) * 100:.0f}%)")
        print(f"\n  {'같은 공간이다 — 더 넓은 표를 밑그림으로 쓸 수 있다' if d[-1] < 1e-6 else '★같은 공간이 아니다 — K 가 바뀌면 노드가 옮겨진다'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())


snapshot._shown = False

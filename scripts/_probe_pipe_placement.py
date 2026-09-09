# -*- coding: utf-8 -*-
"""[배관 위치] 표의 배관이 도면의 «그 자리» 를 가리키는가 — 사용자 지목 추적.

사용자가 배관 11개를 지목했다. 노드로 이으면 한 줄기다::

    9 ─65─ 12 ─40─ 15 ─65─ 18 ─65─ 22 ─40─ 27 ─40─ 32 ─40─ 40 ─40─ 48

그 줄기에서 세 가지가 한꺼번에 보인다.

    ① 관경 역전      65 → 40 → 65. 물길 아래로 갈수록 굵어질 수 없다.
    ② 관경 ↔ 담당 헤드 수  65A·80A 인데 담당 헤드가 1~3개(25A/32A 규모)
    ③ 짧은 조각      P19 0.112 m · P21 0.141 m(elbow-45) · P706 0.15 m

그리고 넷째가 이 셋의 뿌리일 수 있다.

    ④ board 노드쌍이 튄다   대부분 111~205 인데 P722 는 `111–2478`

`edge_ref`(kfp 배관 → board 간선)는 **관경 텍스트를 붙이는 자**다 — `bore.py`
머리말이 「이 표가 없으면 G3 의 관경이 엉뚱한 배관에 붙는다」고 적어 두었다.
노드정리(SSOT)가 배관을 병합하며 id 를 다시 매기므로 그 표는 뒤에 **복구**되는데
(`planar.py`), 그 복구가 어긋나면 관경이 다른 자리의 치수를 물고 온다.

여기서 재는 것:

    · `edge_ref` 가 가리키는 board 좌표 ↔ 그 배관의 실제 표 좌표 (거리)
    · 관경 역전 — 부모(상류)가 자식(하류)보다 얇은 배관
    · 담당 헤드 수 대비 관경이 과대한 배관
    · 짧은 조각 배관

    python scripts/_probe_pipe_placement.py [--k 30] [--key 저장본이름]
"""
from __future__ import annotations

import argparse
import math
import os
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
# 담당 헤드 수로 볼 때 «이 관경이면 과하다» 의 자 — NFPC 별표1 의 상한이 아니라
# 어긋남을 **찾는** 자다(판정이 아니라 계측).
BORE_FOR_LOAD = {1: 25, 2: 32, 3: 40, 5: 50, 10: 65, 30: 80}


def wait(c, sid, limit=40000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _expect_bore(load):
    for n in sorted(BORE_FOR_LOAD):
        if load <= n:
            return BORE_FOR_LOAD[n]
    return 100


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--key", default="")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    if not args.key and not PLAN.is_file():
        print(f"표본 없음: {PLAN}")
        return 0
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        if args.key:
            r = c.post("/api/module-f/reopen", json={"key": args.key})
            j0 = r.get_json() or {}
            if not j0.get("ok"):
                print(f"★이어서 열기 실패 — {str(j0)[:200]}")
                return 1
            sid = j0["sid"]
            if wait(c, sid).get("state") != "done":
                print("★이어서 열기 잡 실패")
                return 1
        else:
            with open(PLAN, "rb") as fh:
                r = c.post("/api/module-f/slot/open",
                           data={"dxf_file": (fh, PLAN.name), "kind": "plan"},
                           content_type="multipart/form-data")
            sid = (r.get_json() or {})["sid"]
            wait(c, sid)
            c.post("/api/module-f/slot/read",
                   json={"sid": sid, "method": "manual"})
            rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
                   .get("recon") or {})
            c.post("/api/module-f/pick/adopt",
                   json={"sid": sid, "materials": True,
                         "heads": {"conf_min": (rec.get("adopt") or {})
                                   .get("conf_min")}})
            wait(c, sid)
            c.post("/api/module-f/pick/commit", json={"sid": sid})
            wait(c, sid)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": args.k})
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {j}")
            return 1

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        d = sess["design"]
        tbl, got = d["tables"], d["got"]
        es = sess["edit"]
        _report(tbl, got, es, args)
    return 0


def _report(tbl, got, es, args):
    edge_ref = got.get("edge_ref") or {}
    origin = got.get("origin_mm")
    board = list(getattr(es.board, "pts", None) or ())
    at = {str(n["label"]): (float(n.get("x", 0) or 0),
                            float(n.get("y", 0) or 0)) for n in tbl.nodes}
    print(f"\n■ 배관 위치 계측 · K={args.k}"
          f" · 배관 {len(tbl.pipes)} · 역참조 {len(edge_ref)}"
          f" · board 점 {len(board)} · origin {origin}")

    # ── ① 역참조가 가리키는 자리 ↔ 그 배관의 실제 자리
    #
    #   board 월드 mm → 표 mm:  (b − origin) + 1000   (`tables._worst_head_node`
    #   가 쓰는 그 식이다 · 1000 은 §T3 재원점 여백)
    def to_table(bxy):
        return ((bxy[0] - float(origin[0])) + 1000.0,
                (bxy[1] - float(origin[1])) + 1000.0)

    far, miss, ok = [], 0, 0
    for p in tbl.pipes:
        pid = str(p.get("label"))
        ref = edge_ref.get(pid)
        a, b = at.get(str(p.get("in"))), at.get(str(p.get("out")))
        if not ref or a is None or b is None or origin is None:
            miss += 1
            continue
        try:
            bi, bj = board[int(ref[0])], board[int(ref[1])]
        except (IndexError, TypeError, ValueError):
            miss += 1
            continue
        ta, tb = to_table(bi), to_table(bj)
        # 어느 쪽으로 대응돼도 되게 두 조합 중 가까운 것을 쓴다.
        d1 = math.dist(ta, a) + math.dist(tb, b)
        d2 = math.dist(ta, b) + math.dist(tb, a)
        gap = min(d1, d2) / 2.0
        if gap > 500.0:
            far.append((pid, round(gap), p.get("dia"), p.get("length"),
                        f"{ref[0]}–{ref[1]}"))
        else:
            ok += 1
    far.sort(key=lambda r: -r[1])
    print("\n  ── ① 역참조가 가리키는 자리 ↔ 배관의 실제 자리")
    print(f"    맞음(500 mm 안) {ok} · **어긋남 {len(far)}** · 역참조 없음 {miss}")
    for r in far[:12]:
        print(f"        {r[0]:8} 어긋남 {r[1]:>7,} mm · {r[2]}A"
              f" · {r[3]} m · board {r[4]}")

    # ── ② 관경 역전 — 상류가 하류보다 얇다
    #
    #   물 흐르는 방향은 표의 in→out 이다(BFS 로 뿌리에서 뻗은 순서).
    nxt: dict = {}
    for p in tbl.pipes:
        nxt.setdefault(str(p.get("in")), []).append(p)
    inv = []
    for p in tbl.pipes:
        up = int(p.get("dia") or 0)
        for q in nxt.get(str(p.get("out")), ()):
            dn = int(q.get("dia") or 0)
            if dn > up:
                inv.append((str(p.get("label")), up, str(q.get("label")), dn))
    print(f"\n  ── ② 관경 역전 (상류 < 하류) **{len(inv)}건**")
    for r in inv[:12]:
        print(f"        {r[0]:8} {r[1]}A  →  {r[2]:8} {r[3]}A")

    # ── ③ 담당 헤드 수 대비 관경
    # 담당 헤드 수는 표 행이 아니라 `tree_loads` 가 권위다(화면도 그것을 쓴다).
    tl = {str(k): int(v) for k, v in (got.get("tree_loads") or {}).items()}
    big = []
    for p in tbl.pipes:
        load = tl.get(str(p.get("label")), 0)
        dia = int(p.get("dia") or 0)
        if load <= 0 or dia <= 0:
            continue
        want = _expect_bore(load)
        if dia >= want * 2:
            big.append((str(p.get("label")), dia, load, want,
                        p.get("dia_src")))
    big.sort(key=lambda r: -(r[1] / max(1, r[3])))
    print(f"\n  ── ③ 담당 헤드 수에 비해 관경이 과대 **{len(big)}건**")
    for r in big[:12]:
        print(f"        {r[0]:8} {r[1]}A · 담당 헤드 {r[2]} (그 부하면 ~{r[3]}A)"
              f" · 근거 {r[4]}")

    # ── ④ 짧은 조각
    short = [(str(p.get("label")), float(p.get("length") or 0),
              p.get("dia"))
             for p in tbl.pipes if 0 < float(p.get("length") or 0) < 0.3]
    short.sort(key=lambda r: r[1])
    print(f"\n  ── ④ 0.3 m 미만 조각 배관 **{len(short)}건**")
    for r in short[:12]:
        print(f"        {r[0]:8} {r[1]:.3f} m · {r[2]}A")
    print(f"\n  관경 근거 분포 {dict(Counter(str(p.get('dia_src')) for p in tbl.pipes))}")

    # ── ⑤ 그 조각들의 정체 — 45° 챔퍼인가
    #
    #   ★한 실제 배관이 «직선 – 챔퍼 – 직선» 으로 잘리면 조각마다 근처 치수
    #     텍스트를 제각기 물어 관경이 튄다. 그것이 ②③ 의 뿌리인지 여기서 본다.
    def ang(p):
        a, b = at.get(str(p.get("in"))), at.get(str(p.get("out")))
        if not a or not b:
            return None
        dx, dy = b[0] - a[0], b[1] - a[1]
        if abs(dx) < 1e-9 and abs(dy) < 1e-9:
            return None
        return math.degrees(math.atan2(dy, dx)) % 180.0

    def bucket(v):
        if v is None:
            return "길이0"
        if v <= 2 or v >= 178:
            return "가로"
        if abs(v - 90) <= 2:
            return "세로"
        if abs(v - 45) <= 3 or abs(v - 135) <= 3:
            return "★45° 대각"
        return "기타"

    short_p = [p for p in tbl.pipes
               if 0 < float(p.get("length") or 0) < 0.3]
    print(f"    조각의 평면 각도 {dict(Counter(bucket(ang(p)) for p in short_p))}")
    long_p = [p for p in tbl.pipes if float(p.get("length") or 0) >= 0.3]
    print(f"    긴 배관의 각도   {dict(Counter(bucket(ang(p)) for p in long_p))}")

    # 조각이 «직선 두 개 사이에» 끼어 있는가 — 그러면 한 배관을 자른 것이다.
    deg: dict = {}
    for p in tbl.pipes:
        deg[str(p.get("in"))] = deg.get(str(p.get("in")), 0) + 1
        deg[str(p.get("out"))] = deg.get(str(p.get("out")), 0) + 1
    inline = sum(1 for p in short_p
                 if deg.get(str(p.get("in")), 0) == 2
                 and deg.get(str(p.get("out")), 0) == 2)
    print(f"    ★조각 {len(short_p)} 중 «양 끝이 차수 2»(한 배관을 자른 것)"
          f" {inline}")

    # 같은 줄기에서 관경이 튀는 자리 — 조각이 낀 곳인가
    by_short = 0
    for a_lab, _u, b_lab, _d in inv:
        pa = next((p for p in tbl.pipes if str(p.get("label")) == a_lab), None)
        pb = next((p for p in tbl.pipes if str(p.get("label")) == b_lab), None)
        if pa is None or pb is None:
            continue
        if (float(pa.get("length") or 0) < 0.3
                or float(pb.get("length") or 0) < 0.3):
            by_short += 1
    print(f"    ★관경 역전 {len(inv)}건 중 «조각이 낀» 것 {by_short}")


if __name__ == "__main__":
    raise SystemExit(main())

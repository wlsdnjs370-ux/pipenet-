# -*- coding: utf-8 -*-
"""[두 화면 선정일치] 평면 보기와 수리계산 표가 **같은 헤드**를 그리는가 — §3.

■ 증상 (사용자 화면 · 같은 도면 · 같은 세션)

      ① 「평면에서 보기」 켬   drawEdit()    ← S.edit.worst  (손질 선정)
      ② 셋 다 끔              drawDesign()  ← 표 5장         (표 선정)

  좌표로 맞대면 27쌍이 회전 없는 한 자로 딱 맞고 **4개만 바꿔치기**돼 있었다.
  기하는 무결하고 다른 것은 «어느 헤드를 골랐는가» 하나뿐이었다.

■ 무엇을 재나

  이 스크립트는 **한 번 돌려 두 값을 다 낸다.** 조치 뒤로는 손질 원본이
  `sess["worst_edit"]` 에 그대로 남으므로, 같은 세션에서

      조치 전 = 화면이 그렸을 집합  = sess["worst_edit"]["heads"]  (손질 원본)
      조치 후 = 화면이 그리는 집합  = sess["worst"]["heads"]       (표 선정)

  둘을 각각 **표의 노즐 좌표**와 1:1 로 맞댄다. 서로 다른 두 실행을 비교하는
  것이 아니라 **같은 실행 안에서** 재는 것이라, 표본이 달라 생기는 흔들림이
  없다. 교체가 일어나지 않은 표본에서는 두 값이 같게 나오고, 그 사실도
  그대로 적는다(없는 차이를 지어내지 않는다).

    python scripts/_probe_two_view_agreement.py [--key 저장본] [--k 10]
                                                [--frac 0.55] [--tag 조치후]
"""
from __future__ import annotations

import argparse
import math
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"

# ★같은 헤드로 볼 거리는 **헤드 반지름에 딸린 값**이다 — 상수로 두면 자가 짧다.
#
#   표의 노즐 절점은 헤드 «중심» 이 아닐 수 있다. 상향식은 배관 위에 티로
#   올라앉으므로 접속 절점이 원 밑을 지나는 관 위에 있고, 그러면 중심에서
#   최대 반지름만큼 떨어진다. 실측 — 대명동(하향식 r42) 최대 33mm ·
#   B1F(상향식) 기준 헤드 150.0mm. 100mm 고정으로 재니 B1F 가 **0/30** 으로
#   나와 「좌표가 어긋났다」로 읽힐 뻔했다. 기하는 무결한데 자가 틀린 것이다.
def near_for(r):
    return float(r) + 60.0        # 반지름 + 붙었다 자(50) + 여유


def wait(c, sid, limit=200000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def match_1to1(want, have):
    """★**1:1** 로 짝짓는다 — 「각자 가장 가까운 것」만 보면 두 헤드가 같은
    노즐을 짚어도 둘 다 «찾음» 이 되어, 빠진 헤드를 놓친다(실측으로 겪었다).

    want = [(x, y, r)] — 문턱을 헤드마다 제 반지름에서 잰다.
    반환: (맞은 거리 목록, 못 찾은 수)
    """
    used, gaps, miss = set(), [], 0
    for (qx, qy, qr) in want:
        q = (qx, qy)
        best, bd = None, 1e18
        for t, p in enumerate(have):
            if t in used:
                continue
            dd = math.dist(p, q)
            if dd < bd:
                best, bd = t, dd
        if best is None or bd > near_for(qr):
            miss += 1
            continue
        used.add(best)
        gaps.append(bd)
    gaps.sort()
    return gaps, miss


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default="")
    ap.add_argument("--k", type=int, default=10)
    ap.add_argument("--frac", type=float, default=0.55)
    ap.add_argument("--tag", default="")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        if args.key:
            j0 = (c.post("/api/module-f/reopen",
                         json={"key": args.key}).get_json() or {})
            if not j0.get("ok"):
                print(f"★이어서 열기 실패 — {str(j0)[:200]}")
                return 1
            sid = j0["sid"]
            if wait(c, sid).get("state") != "done":
                print("★이어서 열기 잡 실패")
                return 1
        else:
            if not PLAN.is_file():
                print(f"표본 없음: {PLAN}")
                return 0
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

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        b = sess["edit"].board
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)

        disks = [(float(d[0]), float(d[1]), float(d[2])) for d in b.disks]
        xs = [p[0] for p in disks]
        ys = [p[1] for p in disks]
        cx, cy = (min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2
        hw = (max(xs) - min(xs)) * args.frac / 2
        hh = (max(ys) - min(ys)) * args.frac / 2
        zone = [cx - hw, cy - hh, cx + hw, cy + hh]
        print(f"\n■ 두 화면 선정일치 계측 {args.tag}"
              f" · 도면 헤드 {len(disks)} · K={args.k}")

        jw = (c.post("/api/module-f/edit/worst",
                     json={"sid": sid, "k": args.k,
                           "zones": [zone]}).get_json() or {})
        if not jw.get("ok"):
            print(f"★최불리 선정 실패 — {str(jw)[:260]}")
            return 1
        edit_heads = [int(i) for i in ((sess.get("worst") or {})
                                       .get("heads") or ())]
        print(f"  [손질] 고른 헤드 {len(edit_heads)}개"
              f" · 영역 안 후보 {(sess.get('worst') or {}).get('candidates')}")

        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {str(j)[:200]}")
            return 1
        d = sess["design"]
        tbl, got = d["tables"], d["got"]
        origin = got.get("origin_mm")
        noz = [str(z.get("in")) for z in tbl.nozzles]
        at = {str(n["label"]): (float(n.get("x", 0) or 0),
                                float(n.get("y", 0) or 0)) for n in tbl.nodes}
        have = [at[n] for n in noz if n in at]

        def to_tbl(p):
            """board mm → 표 좌표. 반지름은 그대로 들고 간다(문턱에 쓴다)."""
            return (p[0] - float(origin[0]) + 1000.0,
                    p[1] - float(origin[1]) + 1000.0,
                    p[2] if len(p) > 2 else 0.0)

        # ── 조치 전 = 손질 원본(교체가 없었으면 그것이 곧 지금 선정이다)
        before = [int(i) for i in ((sess.get("worst_edit") or {})
                                   .get("heads") or ())] or edit_heads
        # ── 조치 후 = **화면이 실제로 그리는** 집합. 좌표를 화면이 받는 그대로
        #    `/edit/state` 에서 읽는다 — 세션을 들여다보면 «화면도 그럴 것이다»
        #    라는 가정이 하나 낀다.
        st2 = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        wv = st2.get("worst") or {}
        shown = [(float(h[0]), float(h[1]), float(h[2]))
                 for h in (wv.get("heads") or ())]

        g0, m0 = match_1to1([to_tbl(disks[i]) for i in before], have)
        g1, m1 = match_1to1([to_tbl(p) for p in shown], have)
        h = (((d.get("got") or {}).get("handoff")) or {})
        print(f"\n  {'항목':<34}{'조치 전':>12}{'조치 후':>12}")
        print("  " + "─" * 58)
        print(f"  {'선정 헤드 수':<32}{len(before):>12}{len(shown):>12}")
        print(f"  {'표 노즐 수':<33}{len(noz):>12}{len(noz):>12}")
        print(f"  {'표와 1:1 로 맞은 헤드':<27}{len(g0):>12}{len(g1):>12}")
        print(f"  {'못 맞은 헤드':<32}{m0:>12}{m1:>12}")
        print(f"  {'최대 오차 (mm)':<30}"
              f"{(f'{g0[-1]:.1f}' if g0 else '—'):>12}"
              f"{(f'{g1[-1]:.1f}' if g1 else '—'):>12}")
        print("  " + "─" * 58)
        print(f"  handoff.filled          {h.get('filled')}")
        print(f"  handoff.swapped_out     {len(h.get('swapped_out') or ())}")
        print(f"  handoff.swapped_in      {len(h.get('swapped_in') or ())}")
        print(f"  marks.swapped_out       "
              f"{((d.get('marks') or {}).get('swapped_out') or {}).get('n', 0)}")

        for r in (h.get("swapped_out") or ())[:8]:
            print(f"    · 빠진 헤드 {r.get('disk')} {r.get('xy')} — "
                  f"{r.get('why_text') or r.get('why') or '사유 미상'}")
        for m in (h.get("messages") or ()):
            print(f"    [문구] {m}")

        # ── 최원 경로가 두 화면에서 같은 줄인가 (§4 기준 2)
        wpath = wv.get("worst_path") or []
        wh = (sess.get("worst") or {}).get("worst_head")
        far_tbl = dict(tbl.meta).get("최원 유하거리 (m)")
        print(f"\n  [최원] 화면 far_m {wv.get('far_m')}"
              f" · 표 meta {far_tbl}"
              f" · 경로 절점 {len(wpath)}")
        if wh is not None and 0 <= int(wh) < len(disks):
            q = to_tbl(disks[int(wh)])[:2]
            near = min(((math.dist(p, q), n) for n, p in at.items()
                        if n in set(noz)), default=(None, None))
            print(f"  [기준 헤드] 화면 disk {wh} → 표 노즐 {near[1]}"
                  + (f" · {near[0]:.1f} mm" if near[0] is not None else " (못 찾음)"))

        ok = (len(g1) == len(shown) == len(noz) and m1 == 0)
        print("\n  " + ("★두 화면이 같은 헤드를 그린다 — 1:1 전부 맞음"
                        if ok else
                        "★아직 어긋난다 — 지시서 §6 대로 여기서 멈추고 보고할 것"))
        if not h.get("filled"):
            print("  (이 표본에서는 교체가 일어나지 않았다 — 그때는 조치 전후가"
                  " 원래 같다. 교체가 있는 표본은 --k 를 올려서 만든다.)")
        return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())

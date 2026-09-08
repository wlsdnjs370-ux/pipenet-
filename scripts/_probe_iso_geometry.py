# -*- coding: utf-8 -*-
"""[위상 ②] 아이소로 «그려진 그림» 이 평면과 같은 모양인가.

앞선 조사(`_probe_design_topology.py`)는 **자료의 위상**이 같다는 것까지 봤다
(갈림 21 · 끝 23 · 사슬 43 · 고리 0 — 양쪽 동일). 그런데 사용자는 여전히
「아이소 변환 시 위상이 깨진다」고 한다. 그렇다면 깨지는 곳은 자료가 아니라
**그려진 그림** 이다.

그래서 여기서는 좌표를 직접 받아 그림을 잰다:

    ① 절점끼리 겹침      — 서로 다른 절점이 같은 자리에 오면 «합쳐진» 것으로 보인다
    ② 선분 교차          — 없던 교차가 생기면 «없는 연결» 로 보인다
    ③ 길이 0 선분        — 사라진 배관
    ④ 절점이 남의 선 위   — T 가 아닌데 T 처럼 보인다

  평면(iso=0)과 아이소(iso=1)를 같은 자로 재서 **아이소에서만 생기는 것**을
  가른다. 아이소 사상 자체는 선형(가역)이라 교차를 새로 만들 수 없다 —
  만들 수 있는 것은 절점마다 다르게 얹히는 **표고 lift** 다. 그 가설도 함께
  확인한다(lift 를 0 으로 두고 다시 잰다).

    python scripts/_probe_iso_geometry.py [도면.dxf]
"""
from __future__ import annotations

import math
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _segs(view):
    at = {str(n["label"]): (float(n["x"]), float(n["y"]))
          for n in (view.get("nodes") or [])}
    out = []
    for p in (view.get("pipes") or []):
        a, b = str(p.get("a")), str(p.get("b"))
        if a in at and b in at:
            out.append((a, b, at[a], at[b]))
    return at, out


def _cross(p1, p2, p3, p4):
    """두 선분이 «끝점을 빼고» 실제로 교차하는가."""
    def d(o, a, b):
        return ((a[0] - o[0]) * (b[1] - o[1])
                - (a[1] - o[1]) * (b[0] - o[0]))
    d1, d2 = d(p3, p4, p1), d(p3, p4, p2)
    d3, d4 = d(p1, p2, p3), d(p1, p2, p4)
    return ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0))


def _pt_seg(px, py, a, b):
    dx, dy = b[0] - a[0], b[1] - a[1]
    L2 = dx * dx + dy * dy
    t = 0.0 if L2 == 0 else max(0.0, min(
        1.0, ((px - a[0]) * dx + (py - a[1]) * dy) / L2))
    return math.hypot(px - (a[0] + t * dx), py - (a[1] + t * dy))


def measure(view, label, eps_ratio=0.004):
    at, segs = _segs(view)
    if not segs:
        print(f"  [{label}] 선분이 없다")
        return {}
    xs = [p[0] for p in at.values()]
    ys = [p[1] for p in at.values()]
    span = max(max(xs) - min(xs), max(ys) - min(ys)) or 1.0
    eps = span * eps_ratio

    # ① 절점 겹침 — 격자로 훑는다
    cell = {}
    dup = 0
    pairs = []
    for lab, (x, y) in at.items():
        key = (int(x // eps), int(y // eps))
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for lab2, (x2, y2) in cell.get((key[0] + dx, key[1] + dy), ()):
                    if math.hypot(x - x2, y - y2) <= eps:
                        dup += 1
                        if len(pairs) < 5:
                            pairs.append((lab2, lab))
        cell.setdefault(key, []).append((lab, (x, y)))

    # ② 교차 · ③ 길이 0
    zero = sum(1 for (_a, _b, p, q) in segs
               if math.hypot(q[0] - p[0], q[1] - p[1]) <= eps * 0.25)
    cx = 0
    show = []
    for i in range(len(segs)):
        a1, b1, p1, p2 = segs[i]
        for j in range(i + 1, len(segs)):
            a2, b2, p3, p4 = segs[j]
            if {a1, b1} & {a2, b2}:
                continue
            if _cross(p1, p2, p3, p4):
                cx += 1
                if len(show) < 5:
                    show.append((f"{a1}-{b1}", f"{a2}-{b2}"))

    # ④ 절점이 남의 선 위에 얹힘
    on = 0
    for lab, (x, y) in at.items():
        for (a, b, p, q) in segs:
            if lab in (a, b):
                continue
            if _pt_seg(x, y, p, q) <= eps * 0.5:
                on += 1
                break

    print(f"  [{label}] 절점 {len(at)} · 선분 {len(segs)} · 폭 {span:.0f}"
          f" (eps {eps:.1f})")
    print(f"      ① 겹친 절점쌍 {dup}{' ' + str(pairs) if pairs else ''}")
    print(f"      ② 교차 {cx}{' ' + str(show) if show else ''}")
    print(f"      ③ 길이 0 선분 {zero}")
    print(f"      ④ 남의 선 위 절점 {on}")
    return {"dup": dup, "cross": cx, "zero": zero, "on": on}


def _detail(view, ev=None):
    """교차가 «왜» 생겼나 — 두 선분의 표고차와 헤드 여부를 함께 본다.

    ★표고는 **표에서** 읽어 넘겨받는다. 미리보기 JSON 의 nodes 는 label·x·y
      만 싣기 때문에 거기서 읽으면 전부 0 으로 보인다 — 한 번 그렇게 읽고
      「표고가 전부 0」이라고 잘못 적었다.
    """
    at, segs = _segs(view)
    meta = {str(n["label"]): n for n in (view.get("nodes") or [])}
    ev = dict(ev or {})
    for lab in meta:
        ev.setdefault(lab, 0.0)
    hd = {lab for lab, n in meta.items() if n.get("head")}
    hits = []
    for i in range(len(segs)):
        a1, b1, p1, p2 = segs[i]
        for j in range(i + 1, len(segs)):
            a2, b2, p3, p4 = segs[j]
            if {a1, b1} & {a2, b2}:
                continue
            if _cross(p1, p2, p3, p4):
                hits.append((a1, b1, a2, b2))
    if not hits:
        return
    n_head = sum(1 for (a1, b1, a2, b2) in hits
                 if {a1, b1} & hd or {a2, b2} & hd)
    print(f"      교차 {len(hits)}건 중 헤드 스텁이 낀 것 {n_head}")
    for (a1, b1, a2, b2) in hits[:6]:
        print(f"        {a1}({ev[a1]:+.2f})-{b1}({ev[b1]:+.2f})"
              f"  ×  {a2}({ev[a2]:+.2f})-{b2}({ev[b2]:+.2f})"
              f"   헤드 {sorted({a1, b1, a2, b2} & hd)}")
    es = sorted({round(v, 3) for v in ev.values()})
    print(f"      표고 종류 {len(es)}: {es[:10]}{' …' if len(es) > 10 else ''}")

    # ★진짜 3차원 교차인가 — 교차점에서 두 배관의 표고가 다르면, 그것은
    #   «위로 지나가는» 실제 배관이다(아이소 그림에서 정상). 표고가 같으면
    #   그림이 만든 가짜 교차다.
    def _at_cross(a1, b1, a2, b2):
        (x1, y1), (x2, y2) = at[a1], at[b1]
        (x3, y3), (x4, y4) = at[a2], at[b2]
        d = (x2 - x1) * (y4 - y3) - (y2 - y1) * (x4 - x3)
        if abs(d) < 1e-12:
            return None
        t = ((x3 - x1) * (y4 - y3) - (y3 - y1) * (x4 - x3)) / d
        u = ((x3 - x1) * (y2 - y1) - (y3 - y1) * (x2 - x1)) / d
        z1 = ev[a1] + (ev[b1] - ev[a1]) * t
        z2 = ev[a2] + (ev[b2] - ev[a2]) * u
        return z1, z2

    real = fake = 0
    for (a1, b1, a2, b2) in hits:
        zz = _at_cross(a1, b1, a2, b2)
        if zz is None:
            continue
        if abs(zz[0] - zz[1]) > 1e-6:
            real += 1
        else:
            fake += 1
    print(f"      교차 {len(hits)} = 진짜 3차원 교차 {real}"
          f" · 표고가 같은 «가짜» {fake}")


def _try_rules(view, stub):
    """헤드 스텁 «방향 규칙» 을 바꿔 보고 교차가 몇으로 떨어지는지 잰다.

    지금 규칙은 «표고 차의 부호» 인데, 이 도면은 표고가 전부 0 이라 부호가 늘
    0 이다 → **모든 헤드가 위로** 선다. 아이소에서 +Y 는 가지배관이 뻗는
    방향이기도 해서, 아래 가지의 헤드 스텁이 위 가지를 가로지른다.
    """
    at, segs = _segs(view)
    meta = {str(n["label"]): n for n in (view.get("nodes") or [])}
    heads = {lab for lab, n in meta.items() if n.get("head")}
    nb = {}
    for (a, b, _p, _q) in segs:
        nb.setdefault(a, []).append(b)
        nb.setdefault(b, []).append(a)
    # 부모 = 헤드에 붙은 단 하나의 이웃
    par = {h: nb[h][0] for h in heads if len(nb.get(h, ())) == 1}

    def count(place):
        pos = dict(at)
        for h, p in par.items():
            pos[h] = place(h, p)
        segs2 = [(a, b, pos[a], pos[b]) for (a, b, _p, _q) in segs]
        n = 0
        for i in range(len(segs2)):
            a1, b1, p1, p2 = segs2[i]
            for j in range(i + 1, len(segs2)):
                a2, b2, p3, p4 = segs2[j]
                if {a1, b1} & {a2, b2}:
                    continue
                if _cross(p1, p2, p3, p4):
                    n += 1
        return n

    def up(h, p):
        return (at[p][0], at[p][1] + stub)

    def away(h, p):
        """부모의 «다른 이웃» 이 있는 쪽을 피해 반대로 세운다."""
        others = [q for q in nb.get(p, ()) if q != h]
        dy = sum(at[q][1] - at[p][1] for q in others) if others else -1.0
        return (at[p][0], at[p][1] + (-stub if dy > 0 else stub))

    def perp(h, p):
        """부모가 뻗는 방향의 직각으로 — 가지와 나란해지지 않게."""
        others = [q for q in nb.get(p, ()) if q != h]
        if not others:
            return up(h, p)
        dx = sum(at[q][0] - at[p][0] for q in others) / len(others)
        dy = sum(at[q][1] - at[p][1] for q in others) / len(others)
        L = math.hypot(dx, dy) or 1.0
        return (at[p][0] - dy / L * stub, at[p][1] + dx / L * stub)

    def half(h, p):
        return (at[p][0], at[p][1] + stub * 0.4)

    print("  [스텁 규칙을 바꿔 보면]")
    for name, fn in (("지금 — 늘 위로", up), ("이웃 반대쪽", away),
                     ("가지 직각", perp), ("길이 40%", half)):
        print(f"      {name:14} 교차 {count(fn)}")


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEF
    if not dxf.is_file():
        print(f"표본 없음: {dxf}")
        return 0
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
        rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
               .get("recon") or {})
        ad = rec.get("adopt") or {}
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True,
                     "heads": {"conf_min": ad.get("conf_min")}})
        wait(c, sid)
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        groups = sorted((g.get("segs") or [] for g in st["body_groups"]),
                        key=len, reverse=True)
        seg = groups[0]
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": seg[0], "y": seg[1]})
        wait(c, sid)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 30})
        c.post("/api/module-f/design/build", json={"sid": sid})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {j}")
            return 1

        print(f"\n■ {dxf.name}")
        got = {}
        for iso in (0, 1):
            pv = c.get(f"/api/module-f/design/preview?sid={sid}"
                       f"&iso={iso}").get_json()
            view = pv.get("view") or {}
            got[iso] = measure(view, "평면(iso=0)" if not iso else "아이소(iso=1)")
            if iso:
                stood = pv.get("stood") or {}
                print(f"      lift {stood.get('lift')} · e_ref"
                      f" {stood.get('e_ref')} · stood keys"
                      f" {sorted(stood)[:8]}")
                trows = (pv.get("tables") or {}).get("nodes") or []
                ev = {str(r.get("label")): float(r.get("elevation") or 0.0)
                      for r in trows}
                _detail(view, ev)
                _try_rules(view, float(stood.get("stub") or 0.0))

        print("\n  [아이소에서만 생긴 것]")
        for k, name in (("dup", "겹친 절점쌍"), ("cross", "교차"),
                        ("zero", "길이 0"), ("on", "남의 선 위 절점")):
            a, b = got[0].get(k, 0), got[1].get(k, 0)
            mark = " ★" if b > a else ""
            print(f"    {name:14} 평면 {a:5} → 아이소 {b:5}{mark}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# -*- coding: utf-8 -*-
"""[손질정본 §3·§4] 손질이 정한 것이 수리계산에 그대로 왔는가 — 불변량으로 잰다.

지시서 `ModuleF_손질정본_지시서.md`.

  좌표가 아니라 «모양의 성질» 로 재야 격자 스냅·노드정리 같은 **정당한 변형**과
  진짜 훼손이 갈린다. 절점 수·간선 수는 달라도 되고(일직선 중간 절점 병합 ·
  세로 구간 추가), 아래 다섯은 같아야 한다:

      헤드 수 · b0(연결 덩어리) · b1(독립 고리) · 말단(차수 1) · 티/크로스(3/4)

  ★세로 구간(헤드 접속관·가지 상승)은 평면에서 길이 0 이다. **버리면 안 된다** —
    버리면 헤드가 망에서 떨어져 b0 와 말단이 부풀어 멀쩡한 망도 훼손으로 읽힌다
    (실측으로 한 번 그렇게 잘못 셌다: b0 1 → 31 · 말단 +42). 두 절점을 **합친다.**

  검산: 말단 = 2 + 티 + 2·크로스 − 2·b1 — 양쪽에서 성립하면 세는 과정이 옳다.

    python scripts/_probe_carryover.py [--k 30] [--zones N] [--key 저장본]
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


def wait(c, sid, limit=200000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def topo(edges) -> dict:
    """간선 (a,b) 목록 → 불변량. 이름은 무엇이든 상관없다."""
    deg = defaultdict(int)
    adj = defaultdict(set)
    ue = {(a, b) if a <= b else (b, a) for a, b in edges if a != b}
    for a, b in ue:
        deg[a] += 1; deg[b] += 1
        adj[a].add(b); adj[b].add(a)
    seen, b0 = set(), 0
    for n in list(adj):
        if n in seen:
            continue
        b0 += 1
        q = deque([n]); seen.add(n)
        while q:
            u = q.popleft()
            for v in adj[u]:
                if v not in seen:
                    seen.add(v); q.append(v)
    V, E = len(deg), len(ue)
    return {"V": V, "E": E, "b0": b0, "b1": E - V + b0,
            "말단": sum(1 for d in deg.values() if d == 1),
            "티": sum(1 for d in deg.values() if d == 3),
            "크로스": sum(1 for d in deg.values() if d == 4)}


def check_euler(t) -> bool:
    """말단 = 2 + 티 + 2·크로스 − 2·b1 (차수 5 이상이 없을 때)."""
    return t["말단"] == 2 + t["티"] + 2 * t["크로스"] - 2 * t["b1"]


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

    ctx = (isolated_workdir(prefix="carry_") if not args.key
           else _Null())
    with ctx, srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        if args.key:
            j0 = c.post("/api/module-f/reopen",
                        json={"key": args.key}).get_json() or {}
            if not j0.get("ok"):
                print(f"★이어서 열기 실패 — {str(j0)[:200]}"); return 1
            sid = j0["sid"]
            if wait(c, sid).get("state") != "done":
                print("★열기 잡 실패"); return 1
        else:
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

        body = {"sid": sid, "k": args.k}
        if args.zones:
            ds = [(float(d[0]), float(d[1])) for d in b.disks]
            xs = [p[0] for p in ds]; ys = [p[1] for p in ds]
            lo, hi = min(xs), max(xs)
            step = (hi - lo) / args.zones
            body["zones"] = [[lo + i * step - 1, min(ys) - 1,
                              lo + (i + 1) * step + 1, max(ys) + 1]
                             for i in range(args.zones)]
        t0 = time.perf_counter()
        jw = c.post("/api/module-f/edit/worst", json=body).get_json()
        t_sel = time.perf_counter() - t0
        if not jw.get("ok"):
            print(f"★최불리 실패 — {str(jw)[:260]}"); return 1
        s = jw["summary"]
        na = s.get("not_attachable") or {}
        w = sess["worst"]
        picked = [int(i) for i in (w.get("heads") or ())]

        print(f"\n■ 손질 → 수리계산 인계 계측 · K={args.k}"
              f" · 영역 {args.zones or '없음'}"
              f" · {args.key or '대명동'}")
        print(f"\n[후보]     영역 안 헤드 {na.get('n', 0) + s.get('candidates', 0)}개"
              f" → 붙는 헤드 {s.get('candidates')}개 (뺀 것 {na.get('n', 0)}개)")
        print(f"[사유]     " + (" · ".join(f"{k2} {v}" for k2, v in
                                          (na.get('by_why') or {}).items())
                                or "(없음)"))

        t0 = time.perf_counter()
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        if wait(c, sid).get("state") != "done":
            print("★표 확정 실패"); return 1
        t_tbl = time.perf_counter() - t0
        d = sess["design"]
        got, tbl = d["got"], d["tables"]
        h = got.get("handoff") or {}
        origin = got.get("origin_mm") or (0.0, 0.0)
        ox, oy = float(origin[0]) - 1000.0, float(origin[1]) - 1000.0
        at = {str(n.get("label")): (float(n.get("x") or 0) + ox,
                                    float(n.get("y") or 0) + oy)
              for n in tbl.nodes}
        noz = [at[str(z.get("in"))] for z in tbl.nozzles
               if str(z.get("in")) in at]

        used, gaps, miss = set(), [], 0
        for i in picked:
            q = (float(b.disks[i][0]), float(b.disks[i][1]))
            r = float(b.disks[i][2])
            best, bd = None, 1e18
            for t, p in enumerate(noz):
                if t in used:
                    continue
                dd = math.dist(p, q)
                if dd < bd:
                    best, bd = t, dd
            if best is None or bd > r + 60.0:
                miss += 1; continue
            used.add(best); gaps.append(bd)
        gaps.sort()
        print(f"[선정]     K={args.k} · 손질이 고른 {len(picked)}개"
              f" · 표 노즐 {len(noz)}개 · 좌표 1:1 {len(gaps)}/{len(picked)}"
              + (f" (최대 {gaps[-1]:.1f} mm)" if gaps else "")
              + (f" · 표에만 {len(noz) - len(used)}" if len(noz) != len(used) else ""))
        print(f"[백필]     filled={h.get('filled', 0)}"
              + ("" if not h.get("filled") else "  ★0이어야 한다"))

        # ── §3 불변량 — 손질 corridor ↔ 표
        cor0 = [(int(a), int(c2)) for (a, c2) in (w.get("loads") or {})]
        # ★길이 0 구간 합치기는 **양쪽에 똑같이** 대야 한다.
        #   표에서만 합치고 손질에서는 안 합치면, 한 자리에 board 절점이 둘인
        #   곳(수직 접속·같은 점의 서로 다른 id)이 손질에서는 「티 둘」로,
        #   표에서는 「크로스 하나」로 세어져 멀쩡한 망이 어긋나 보인다.
        #   실측(B1F): 이 비대칭 하나가 티 10/4 · 크로스 0/4 를 만들었다.
        cpar: dict = {}

        def cfind(x):
            cpar.setdefault(x, x)
            while cpar[x] != x:
                cpar[x] = cpar[cpar[x]]; x = cpar[x]
            return x

        n_zero = 0
        for a, z in cor0:
            if math.dist(b.pts[a][:2], b.pts[z][:2]) <= 1.0:
                n_zero += 1
                ra, rz = cfind(a), cfind(z)
                if ra != rz:
                    cpar[ra] = rz
        cor = [(cfind(a), cfind(z)) for a, z in cor0
               if cfind(a) != cfind(z)]
        t_edit = topo(cor)
        parent: dict = {}

        def find(x):
            parent.setdefault(x, x)
            while parent[x] != x:
                parent[x] = parent[parent[x]]; x = parent[x]
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
                vert += 1; union(a, z)
        flat = []
        for pr in tbl.pipes:
            ra, rz = find(str(pr.get("in"))), find(str(pr.get("out")))
            if ra != rz:
                flat.append((ra, rz))
        t_tblv = topo(flat)

        # ★두 물건을 같은 물건으로 만들어 놓고 재야 한다.
        #   손질 corridor 는 「급수원 → 헤드 경로의 합집합」이라 **스텁이 없다**.
        #   표는 `prune_dead_pipes` 가 «끝배관1단보호» 로 마지막 헤드 너머 한
        #   마디를 일부러 남긴다(실측 B1F: 11개). 그 스텁이 옆으로 갈라지면
        #   표에만 티 1 · 말단 1 이 더 생긴다 — 훼손이 아니라 규칙이다.
        #   그래서 **헤드도 급수원도 아닌 말단**을 걷어 낸 뒤 한 번 더 잰다.
        keep_lbl = {find(str(z.get("in"))) for z in tbl.nozzles}
        for n in tbl.nodes:
            if str(n.get("io_node") or "").lower().startswith("in"):
                keep_lbl.add(find(str(n.get("label"))))
        flat2, n_stub = list(flat), 0
        while True:
            dg: dict = {}
            for a, z in flat2:
                dg[a] = dg.get(a, 0) + 1
                dg[z] = dg.get(z, 0) + 1
            cut = {n for n, d2 in dg.items() if d2 == 1 and n not in keep_lbl}
            if not cut:
                break
            n_stub += len(cut)
            flat2 = [(a, z) for a, z in flat2 if a not in cut and z not in cut]
        t_tbl2 = topo(flat2)

        names = ["b0", "b1", "말단", "티", "크로스"]
        same = all(t_edit[n] == t_tbl2[n] for n in names)
        heads_same = len(picked) == len(noz)
        print(f"[불변량]   헤드 {len(picked)}/{len(noz)}"
              + "".join(f" · {n} {t_edit[n]}/{t_tbl2[n]}" for n in names))
        if any(t_tblv[n] != t_tbl2[n] for n in names):
            print(f"           표 원본(보호스텁 {n_stub}개 포함, 참고)"
                  + "".join(f" · {n} {t_tblv[n]}" for n in names))
        print(f"           V·E (달라도 된다) {t_edit['V']}/{t_tblv['V']}"
              f" · {t_edit['E']}/{t_tbl2['E']}"
              f" · 길이0 구간 합침 손질 {n_zero} · 표 {vert}")
        print(f"           검산(말단=2+티+2·크로스−2·b1):"
              f" 손질 {'OK' if check_euler(t_edit) else '★어긋남'}"
              f" · 표 {'OK' if check_euler(t_tbl2) else '★어긋남'}")
        # ── §3 어긋나면 «어디가» 다른지까지 말한다 (찾아 헤매지 않도록)
        if not same:
            def _branch(edges2, xy):
                d2: dict = {}
                for a2, b2 in {(x, y) if x <= y else (y, x)
                               for x, y in edges2 if x != y}:
                    d2[a2] = d2.get(a2, 0) + 1
                    d2[b2] = d2.get(b2, 0) + 1
                return sorted(((xy(n), deg) for n, deg in d2.items()
                               if deg >= 3), key=lambda t: t[0])
            be = _branch(cor, lambda n: (round(float(b.pts[n][0]), 1),
                                         round(float(b.pts[n][1]), 1)))
            bt = _branch(flat2, lambda n: (round(at[n][0], 1), round(at[n][1], 1))
                         if n in at else (0.0, 0.0))
            print(f"[분기점]   손질 {len(be)}곳 · 표 {len(bt)}곳"
                  f" — 자리로 맞춰 봅니다 (300mm 이내를 같은 곳으로)")
            un_t = list(bt)
            for p, dg in be:
                near = min(un_t, key=lambda q: math.dist(p, q[0],),
                           default=None)
                if near and math.dist(p, near[0]) <= 300.0:
                    un_t.remove(near)
                    mark = "" if dg == near[1] else f"  ★차수 {dg}→{near[1]}"
                    if mark:
                        print(f"    ({p[0]:.0f}, {p[1]:.0f}){mark}")
                else:
                    print(f"    ({p[0]:.0f}, {p[1]:.0f}) 차수 {dg}"
                          f"  ★표에 같은 자리 분기점이 없다")
            for q, dg in un_t:
                print(f"    ({q[0]:.0f}, {q[1]:.0f}) 차수 {dg}"
                      f"  ★손질에 없던 분기점이 표에 있다")
            # 손질 분기점끼리 붙어 있으면 어떤 뭉침에서도 하나로 접힌다
            close = [(p, q) for i2, (p, _) in enumerate(be)
                     for (q, _) in be[i2 + 1:] if math.dist(p, q) <= 300.0]
            print(f"           손질 분기점끼리 300mm 안에 붙은 짝"
                  f" {len(close)}개" + (" — 뭉치면 크로스가 된다"
                                        if close else ""))

        print(f"[시간]     최불리 {t_sel:.1f}s · 표 확정 {t_tbl:.1f}s")

        ok = (heads_same and same and miss == 0
              and not h.get("filled") and len(noz) == len(used))
        print("\n  " + ("★수용 기준 1·2·5 가 선다" if ok else
                        "★★아직 안 선다 — 위 숫자가 어디가 다른지 말한다"))
        return 0 if ok else 2


class _Null:
    def __enter__(self): return None
    def __exit__(self, *a): return False


if __name__ == "__main__":
    raise SystemExit(main())

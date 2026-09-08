# -*- coding: utf-8 -*-
"""[위상] 평면에서 본 배관망과 아이소에서 본 배관망이 **같은 위상인가**.

사용자 지적: 「평면에서 보기에서 본 배관망이랑 30° 아이소매트릭에서 본
배관망이랑 위상이 조금 다른데? 형태를 옮기는 과정에서 뭔가가 훼손된 것 같다.」

두 화면은 서로 다른 자료를 그린다:

    평면  — 손질판(board)의 절점·간선. 최불리 corridor 가 굵게.
    아이소 — 「표 확정」이 만든 **설계 좌표**(display_tables 의 view).

그래서 «모양» 이 다른 것은 당연하다(스키매틱 배치다). 여기서 재는 것은 모양이
아니라 **위상**이다:

    ① corridor 간선 수 ↔ 설계 배관 수(수직·스텁 빼고)
    ② 차수 분포(1·2·3·4 차 절점 수)
    ③ **고리(cycle) 수** — 평면에 고리가 있는데 설계가 나무면 위상이 깨진 것
    ④ edge_ref 로 짝지어지지 않은 배관 · corridor 간선

    python scripts/_probe_design_topology.py [도면.dxf]
"""
from __future__ import annotations

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


def _components(adj):
    seen, comps = set(), 0
    for n in adj:
        if n in seen:
            continue
        comps += 1
        stack = [n]
        seen.add(n)
        while stack:
            u = stack.pop()
            for v in adj[u]:
                if v not in seen:
                    seen.add(v)
                    stack.append(v)
    return comps


def _shape(adj, name):
    """절점·간선·고리·차수 분포 한 줄."""
    v = len(adj)
    e = sum(len(s) for s in adj.values()) // 2
    c = _components(adj) if v else 0
    deg = {}
    for n, s in adj.items():
        deg[len(s)] = deg.get(len(s), 0) + 1
    cyc = e - v + c
    print(f"    {name:14} 절점 {v:4} · 간선 {e:4} · 조각 {c:2} · 고리 {cyc:3}"
          f" · 차수 {dict(sorted(deg.items()))}")
    return {"v": v, "e": e, "c": c, "cycles": cyc, "deg": deg}


def _skeleton(adj):
    """차수 2 절점을 접어 «뼈대» 만 남긴다 — 위상 비교의 자.

    일직선 중간 절점 병합(SSOT)이나 스텁 쪼개기는 위상을 바꾸지 않는데, 절점
    수만 견주면 그것들이 «달라진 것» 으로 잡힌다. 갈림(3차 이상)과 끝(1차)만
    남기면 두 그래프가 정말 같은 모양인지 보인다.
    """
    keep = {n for n, s in adj.items() if len(s) != 2}
    if not keep:
        return {"j": 0, "leaf": 0, "chains": 0, "cycles": 0}
    chains = 0
    seen = set()
    for a in keep:
        for nb in adj[a]:
            prev, cur = a, nb
            while cur not in keep:
                nxt = [x for x in adj[cur] if x != prev]
                if not nxt:
                    break
                prev, cur = cur, nxt[0]
            # 같은 두 끝 사이에 사슬이 여럿일 수 있다(고리) — 방향까지 세고 2로 나눈다.
            seen.add((a, nb))
            chains += 1
    chains //= 2
    j = sum(1 for n in keep if len(adj[n]) >= 3)
    leaf = sum(1 for n in keep if len(adj[n]) == 1)
    cyc = chains - len(keep) + _components(adj)
    return {"j": j, "leaf": leaf, "chains": chains, "cycles": cyc}


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
        r = c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 30})
        if not (r.get_json() or {}).get("ok"):
            print(f"★최불리 실패 — {r.get_json()}")
            return 1
        r = c.post("/api/module-f/design/build", json={"sid": sid})
        if not (r.get_json() or {}).get("ok"):
            print(f"★표 확정 실패 — {r.get_json()}")
            return 1
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {j}")
            return 1

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        d = sess["design"]
        got = d["got"]
        es = sess["edit"]
        board = es.board
        worst = got.get("worst") or {}
        ref = got.get("edge_ref") or {}

        print(f"\n■ {dxf.name}")
        # ── ① 평면 쪽 — corridor(최불리 망) 간선
        cor = set()
        for e in (worst.get("edges") or ()):
            try:
                i, j2 = int(e[0]), int(e[1])
            except (TypeError, ValueError, IndexError):
                continue
            cor.add((min(i, j2), max(i, j2)))
        adj_b = {}
        for (i, j2) in cor:
            adj_b.setdefault(i, set()).add(j2)
            adj_b.setdefault(j2, set()).add(i)
        print("  [평면] 손질판 corridor")
        b = _shape(adj_b, "corridor")
        print(f"    board 전체 절점 {len(board.pts)} · 간선 {len(board.edges)}")

        # ── ② 아이소 쪽 — 설계 표(view)의 배관
        pv = c.get(f"/api/module-f/design/preview?sid={sid}").get_json()
        view = pv.get("view") or {}
        pipes = view.get("pipes") or []
        nodes = view.get("nodes") or []
        adj_d = {}
        # 화면이 쓰는 키와 같은 것을 쓴다(a/b). 표 자체는 in/out 이지만
        # 미리보기 JSON 은 화면용으로 a·b 로 옮겨 담는다 — 여기서 in/out 을
        # 읽으면 간선이 하나도 안 잡힌다(처음에 그렇게 틀렸다).
        for row in pipes:
            a = str(row.get("a", row.get("in")))
            bb = str(row.get("b", row.get("out")))
            adj_d.setdefault(a, set()).add(bb)
            adj_d.setdefault(bb, set()).add(a)
        print("  [아이소] 설계 표")
        dsh = _shape(adj_d, "design view")
        print(f"    노즐 {len(view.get('nozzles') or [])}"
              f" · 절점행 {len(nodes)}")

        # ── ③ 짝짓기 — 어느 배관이 어느 corridor 간선에서 왔나
        mapped = set()
        for pid, e in ref.items():
            try:
                i, j2 = int(e[0]), int(e[1])
            except (TypeError, ValueError, IndexError):
                continue
            mapped.add((min(i, j2), max(i, j2)))
        print("  [짝짓기]")
        print(f"    설계 배관 {len(pipes)} · edge_ref 로 도면 간선에 붙은 것 "
              f"{len(ref)} · 나머지(수직·스텁) {len(pipes) - len(ref)}")
        print(f"    corridor 간선 {len(cor)} · 그 중 설계로 옮겨진 것 "
              f"{len(cor & mapped)} · **옮겨지지 못한 것 {len(cor - mapped)}**")
        print(f"    설계가 가리켰는데 corridor 에 없는 간선 {len(mapped - cor)}")

        # ── ③-b 뼈대 — 일직선 병합·스텁을 접고 «갈림과 끝» 만 견준다
        adj_m = {}
        for (i, j2) in mapped:
            adj_m.setdefault(i, set()).add(j2)
            adj_m.setdefault(j2, set()).add(i)
        sk_b = _skeleton(adj_b)
        sk_m = _skeleton(adj_m)
        sk_d = _skeleton(adj_d)
        print("  [뼈대] 차수 2 를 접고 갈림·끝만")
        print(f"    corridor      갈림 {sk_b['j']:3} · 끝 {sk_b['leaf']:3}"
              f" · 사슬 {sk_b['chains']:3} · 고리 {sk_b['cycles']:3}")
        print(f"    설계가 쓴 도면 갈림 {sk_m['j']:3} · 끝 {sk_m['leaf']:3}"
              f" · 사슬 {sk_m['chains']:3} · 고리 {sk_m['cycles']:3}")
        print(f"    설계 표       갈림 {sk_d['j']:3} · 끝 {sk_d['leaf']:3}"
              f" · 사슬 {sk_d['chains']:3} · 고리 {sk_d['cycles']:3}")

        # ── ③-c 사슬 단위 대조 — «설계가 가리킨 도면 간선» 이 corridor 의
        #        어느 사슬에 앉는가. 일직선 병합 때문에 간선 수는 달라도,
        #        사슬은 하나도 빠지거나 남으면 안 된다.
        chain_of = {}
        keepb = {n for n, sset in adj_b.items() if len(sset) != 2}
        cid = 0
        for a in keepb:
            for nb in adj_b[a]:
                prev, cur = a, nb
                path = [(min(a, nb), max(a, nb))]
                while cur not in keepb:
                    nxt = [x for x in adj_b[cur] if x != prev]
                    if not nxt:
                        break
                    path.append((min(cur, nxt[0]), max(cur, nxt[0])))
                    prev, cur = cur, nxt[0]
                if any(e in chain_of for e in path):
                    continue
                cid += 1
                for e in path:
                    chain_of[e] = cid
        hit = {}
        orphan = 0
        for e in mapped:
            k = chain_of.get(e)
            if k is None:
                orphan += 1
            else:
                hit[k] = hit.get(k, 0) + 1
        print("  [사슬 대조]")
        print(f"    corridor 사슬 {cid} · 설계가 앉은 사슬 {len(hit)}"
              f" · **빈 사슬 {cid - len(hit)}**")
        print(f"    corridor 사슬 어디에도 없는 설계 간선 {orphan}")

        # ── ③-d 얼마나 «다른» 길인가 — 떨어진 거리로 잰다.
        #        평행선 중복(같은 자리 다른 선)이면 거리가 거의 0 이고,
        #        정말 다른 길로 돌았으면 미터 단위로 벌어진다.
        import math
        pts = board.pts

        def _mid(e):
            (x1, y1), (x2, y2) = pts[e[0]], pts[e[1]]
            return ((x1 + x2) / 2.0, (y1 + y2) / 2.0)

        def _seg_d(px, py, e):
            (x1, y1), (x2, y2) = pts[e[0]], pts[e[1]]
            dx, dy = x2 - x1, y2 - y1
            L2 = dx * dx + dy * dy
            t = 0.0 if L2 == 0 else max(0.0, min(
                1.0, ((px - x1) * dx + (py - y1) * dy) / L2))
            return math.hypot(px - (x1 + t * dx), py - (y1 + t * dy))

        cor_list = list(cor)
        ds = []
        for e in sorted(mapped - cor):
            mx, my = _mid(e)
            ds.append(min(_seg_d(mx, my, ce) for ce in cor_list))
        print("  [얼마나 다른 길인가]")
        if ds:
            ds.sort()
            near = sum(1 for v in ds if v <= 50.0)
            print(f"    corridor 밖 설계 간선 {len(ds)} · corridor 까지 거리"
                  f" 중앙 {ds[len(ds) // 2]:.0f}mm · 최대 {ds[-1]:.0f}mm")
            print(f"    그 중 50mm 이내(사실상 같은 자리) {near}"
                  f" · 1m 넘게 떨어진 것 {sum(1 for v in ds if v > 1000)}")
        empty_len = 0.0
        for e, k in chain_of.items():
            if k not in hit:
                (x1, y1), (x2, y2) = pts[e[0]], pts[e[1]]
                empty_len += math.hypot(x2 - x1, y2 - y1)
        print(f"    설계가 안 앉은 corridor 사슬 {cid - len(hit)}개 ·"
              f" 그 길이 합 {empty_len / 1000:.1f} m")

        # ── ③-e 짝짓기가 «쓰이는» 자리 — 담당 헤드 수(굵기)와 관경 덮기.
        #        edge_ref 가 corridor 밖 간선을 가리키면 loads 조회가 빗나가
        #        그 배관만 굵기 0 으로 그려진다. 위상은 같아도 «다르게 보이는»
        #        원인이 될 수 있어 함께 센다.
        loads = worst.get("loads") or {}
        miss = 0
        for pid, e in ref.items():
            try:
                i, j2 = int(e[0]), int(e[1])
            except (TypeError, ValueError, IndexError):
                continue
            if (min(i, j2), max(i, j2)) not in loads:
                miss += 1
        tree = got.get("tree_loads") or {}
        rescue = 0
        for pid, e in ref.items():
            try:
                i, j2 = int(e[0]), int(e[1])
            except (TypeError, ValueError, IndexError):
                continue
            if (min(i, j2), max(i, j2)) not in loads and str(pid) in {
                    str(k) for k in tree}:
                rescue += 1
        print("  [굵기 근거]")
        print(f"    edge_ref {len(ref)} · 그 중 담당 헤드 수를 못 찾은 것 {miss}"
              f" ({miss * 100 // max(1, len(ref))}%)")
        print(f"    tree_loads 가 대신 아는 것 {rescue} / {miss}"
              f" · tree_loads 항목 {len(tree)}")
        # 화면이 실제로 받는 값 — 굵기 0 으로 그려지는 배관이 몇인가.
        zero = sum(1 for r in pipes if not (r.get("load") or 0))
        print(f"    미리보기가 굵기 0 으로 준 배관 {zero} / {len(pipes)}")

        # ── ④ 고리 — 위상이 깨지는 가장 흔한 자리
        print("  [고리]")
        print(f"    평면 corridor 고리 {b['cycles']} · 설계 고리 {dsh['cycles']}")
        if b["cycles"] and not dsh["cycles"]:
            print("    ★평면에 고리가 있는데 설계는 나무다 — 위상이 깨졌다.")
        elif b["cycles"] != dsh["cycles"]:
            print("    ★고리 수가 다르다.")
        else:
            print("    고리 수는 같다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# -*- coding: utf-8 -*-
"""[위상 손실] 모듈 F 통합 경계 계측 — 지시서 §2.

`ModuleF_위상손실_수정지시서.md` 의 0·2·4·7 단계에서 **같은 형식**으로 돌린다.
단계 간 비교가 이 작업의 근거이므로 출력 형식을 바꾸지 않는다.

    python scripts/_f_topology_probe.py [--tag 0단계]

대상 두 벌:

    A. 대명동 (평면도 + 계통도)            — 기계실 없음. 알려진 정상 사례(기준선)
    B. 대명동 (평면도 + 계통도 + 기계실)   — 기계실이 붙는 경우

★두 벌 모두 같은 도면 세트를 쓴다(`routes/제출용[최종]/`). 사용자가 화면에서
  전 공정을 태우는 그 세 장이고, 이 저장소에서 기계실까지 붙는 유일한 완전한
  세트다. 다른 세트를 쓰려면 인자로 준다.
"""
from __future__ import annotations

import argparse
import math
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

D = ROOT / "routes" / "제출용[최종]"
PLAN = D / "1. 입력도면 대명동 단위세대 평면도.dxf"
SYS = D / "1. 입력도면 대명동 단위세대 계통도.dxf"
MR = D / "1. 입력도면 대명동 단위세대 기계실.dxf"


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _design_tables(c):
    """평면도 → 설계 표. 화면이 밟는 순서 그대로(찍기→조립→앵커→최불리→확정)."""
    with open(PLAN, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, PLAN.name), "kind": "plan"},
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
    seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                 key=len, reverse=True)[0]
    c.post("/api/module-f/edit/anchor-click",
           json={"sid": sid, "x": seg[0], "y": seg[1]})
    wait(c, sid)
    c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 30})
    c.post("/api/module-f/design/build", json={"sid": sid})
    j = wait(c, sid)
    if j.get("state") != "done":
        raise SystemExit(f"★표 확정 실패 — {j}")
    from routes.module_f.jobs import _sess
    return _sess(sid)["design"]["tables"], _sess(sid).get("method")


def _sub(path, kind):
    """계통도·기계실 추출 — 배관 위 두 점(한 계통의 아래끝·위끝)을 찍는다."""
    from remote30_prototype import _auto_pipe_layer_filter, build_system_graph
    from routes.module_f.subdrawing import (extract_machineroom, extract_system,
                                            parse_subdrawing)
    ents, _ = parse_subdrawing(path)
    auto = sorted(_auto_pipe_layer_filter(ents))
    g, _e, _s = build_system_graph(ents, layer_filter={auto[0]},
                                   force_connect=True)
    ns = sorted(g, key=lambda n: n[1])
    a = (ns[0][0] + 120.0, ns[0][1] - 90.0)
    b = (ns[-1][0] - 80.0, ns[-1][1] + 140.0)
    if kind == "system":
        return extract_system(ents, a, b, layer_filter={auto[0]})
    got = extract_machineroom(ents, a, b, layer_filter={auto[0]})
    got["conn_xy"] = [b[0], b[1]]
    return got


def _components(nodes, pipes) -> list[int]:
    adj = {str(n.get("label")): set() for n in nodes}
    for p in pipes:
        a, b = str(p.get("in")), str(p.get("out"))
        if a in adj and b in adj:
            adj[a].add(b)
            adj[b].add(a)
    seen, sizes = set(), []
    for n in adj:
        if n in seen:
            continue
        stack, size = [n], 0
        seen.add(n)
        while stack:
            u = stack.pop()
            size += 1
            for v in adj[u]:
                if v not in seen:
                    seen.add(v)
                    stack.append(v)
        sizes.append(size)
    return sorted(sizes, reverse=True)


def measure(name, tbl, method, riser, machineroom):
    from routes.module_f.merge import (label_offset_for, merge_network,
                                       riser_tables_from, to_head_tables)
    got = merge_network(tbl, riser=riser, machineroom=machineroom,
                        mode="lsp_gravity", method=method)
    c = got["combined"]
    if c is None:
        print(f"\n=== {name}: 결합망 없음(계통도 없음)")
        return
    nodes, pipes = list(c.nodes), list(c.pipes)

    xs = [float(n.get("x", 0) or 0) for n in nodes]
    ys = [float(n.get("y", 0) or 0) for n in nodes]
    span_x, span_y = max(xs) - min(xs), max(ys) - min(ys)
    longest = max(span_x, span_y)
    scale = (3000.0 / longest) if longest > 1e-9 else 1.0

    comps = _components(nodes, pipes)
    q = sum(1 for p in pipes
            if str(p.get("in")) == "?" or str(p.get("out")) == "?")
    origin = sum(1 for n in nodes
                 if abs(float(n.get("x", 0) or 0)) < 1e-9
                 and abs(float(n.get("y", 0) or 0)) < 1e-9)

    import re
    labs = {str(p.get("label")) for p in pipes}
    renamed = [x for x in labs if re.search(r"_\d+$", x)
               and re.sub(r"_\d+$", "", x) in labs]
    orphan_f = [f for f in (c.fittings or ())
                if str(f.get("pipe")) not in labs]
    orphan_e = [e for e in (c.equipment or ())
                if e.get("pipe") and str(e.get("pipe")) not in labs]

    # 두 AV — stitch 밖에서는 같은 공개 함수로 다시 세워 견준다(D5 전 임시).
    ht = to_head_tables(tbl, offset=label_offset_for(method))
    rt = riser_tables_from(riser)
    mr_labels = []
    if machineroom:
        from remote30_full_network import prepend_machine_room_to_riser
        mr_labels = [str(n.get("label"))
                     for n in (machineroom.get("nodes") or ())]
        rt, ok = prepend_machine_room_to_riser(
            machineroom, rt, at_bottom=False, source_drop_below_lowest_m=0.0)
        if not ok:
            mr_labels = []
    av = str(rt.av_node_label)
    r_av = next((n for n in rt.nodes if str(n.get("label")) == av), None)
    h_av = next((n for n in ht.nodes if str(n.get("label")) == av), None)
    av_d = av_dz = None
    if r_av and h_av:
        av_d = math.dist((float(r_av["x"]), float(r_av["y"])),
                         (float(h_av["x"]), float(h_av["y"])))
        av_dz = (float(r_av.get("elevation", 0) or 0)
                 - float(h_av.get("elevation", 0) or 0))

    # 기계실 노드가 «원 좌표 그대로» 남았는가 — D1 의 직접 증거.
    raw = {str(n.get("label")): (float(n.get("x", 0) or 0),
                                 float(n.get("y", 0) or 0))
           for n in (machineroom.get("nodes") or ())} if machineroom else {}
    at = {str(n.get("label")): (float(n.get("x", 0) or 0),
                                float(n.get("y", 0) or 0)) for n in nodes}
    unmoved = [lab for lab in mr_labels
               if lab in raw and lab in at
               and math.dist(raw[lab], at[lab]) < 1e-6]

    print(f"\n=== {name}")
    print(f"  규모      절점 {len(nodes)} · 배관 {len(pipes)} · 노즐 "
          f"{len(c.nozzles or ())} · 부속 {len(c.fittings or ())} · 기기 "
          f"{len(c.equipment or ())}")
    print(f"  bbox      x [{min(xs):.0f}, {max(xs):.0f}] span {span_x:.0f}"
          f" · y [{min(ys):.0f}, {max(ys):.0f}] span {span_y:.0f}")
    print(f"  emit 배율 3000/{longest:.0f} = {scale:.6f}")
    print(f"  연결성분  {len(comps)}개 {comps[:6]}")
    print(f"  D2 «?» 끝점 배관 {q}")
    print(f"  D3 원점(0,0) 절점 {origin}")
    print(f"  D4 개명 배관 {len(renamed)} {sorted(renamed)[:5]}"
          f" · 고아 부속 {len(orphan_f)} · 고아 기기 {len(orphan_e)}")
    print(f"  두 AV     거리 {('%.1f mm' % av_d) if av_d is not None else '—'}"
          f" · 표고차 {('%.3f m' % av_dz) if av_dz is not None else '—'}")
    print(f"  D1 기계실 라벨 {len(mr_labels)} · 좌표 원본 그대로 {len(unmoved)}"
          f" · plan_edges {len(getattr(c, 'machine_room_plan_edges', ()) or ())}")
    for line in (got.get("steps") or ()):
        print(f"      · {line}")
    return got


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tag", default="")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    for p in (PLAN, SYS, MR):
        if not p.is_file():
            print(f"표본 없음: {p}")
            return 0

    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    print(f"■ 모듈 F 위상 손실 계측 {args.tag}")
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        tbl, method = _design_tables(c)
    riser = _sub(SYS, "system")
    mr = _sub(MR, "machineroom")
    measure("A. 대명동 (평면도+계통도)", tbl, method, riser, None)
    measure("B. 대명동 (평면도+계통도+기계실)", tbl, method, riser, mr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

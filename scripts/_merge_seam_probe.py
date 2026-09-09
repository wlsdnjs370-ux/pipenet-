# -*- coding: utf-8 -*-
"""[이음매 좌표] 세 도면 공통 노드가 한 점에서 만나는가 — 지시서 §4.

`ModuleF_이음매좌표_수정지시서.md` 의 계측 도구. 조치 전/후로 **같은 표**를 낸다.

    · parts 세 목록 크기 · **unclassified**(어디에도 없는 라벨)      ← E1
    · 부위별 bbox 와 서로의 배율                                     ← E3
    · anchor_gap_mm (기대 0) · pump_seam_mm 과 표 length 대조        ← E3
    · layout_status (계통도·기계실 배치가 폴백으로 떨어졌는가)       ← E2
    · bake_combined_iso 뒤 seam_check 두 값 (기대 <1e-6)             ← E3
    · 좌표 거리 ÷ 1000 이 표 length 와 5% 넘게 다른 배관             ← 총괄

대명동 3장(평면 + 계통 + 기계실)을 **펌프 가압**으로 결합한다 — E1 은 그 모드
에서만 난다(`is_pump and pump`).

    python scripts/_merge_seam_probe.py [--tag 조치전]
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
PUMP = {"rated_q_lpm": 800.0, "rated_h_m": 60.0, "count": 2}


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _design_tables(c):
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


def _bbox(nodes):
    xs = [float(n.get("x", 0) or 0) for n in nodes]
    ys = [float(n.get("y", 0) or 0) for n in nodes]
    if not xs:
        return None
    return (min(xs), max(xs), min(ys), max(ys),
            max(max(xs) - min(xs), max(ys) - min(ys)))


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

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        tbl, method = _design_tables(c)
    riser = _sub(SYS, "system")
    mr = _sub(MR, "machineroom")

    from routes.module_f.merge import merge_network
    got = merge_network(tbl, riser=riser, machineroom=mr,
                        mode="hsp_pump", pump=PUMP, method=method,
                        source_drop_m=3.0)
    cb = got["combined"]
    nodes = list(cb.nodes)
    pipes = list(cb.pipes)
    parts = got.get("parts") or {}
    of = {}
    for kind in ("system", "machineroom", "plan"):
        for lab in (parts.get(kind) or ()):
            of[str(lab)] = kind

    print(f"\n■ 이음매 좌표 계측 {args.tag}  (펌프 가압 · 수원 낙차 3.0 m)")
    print(f"  parts  평면 {len(parts.get('plan') or ())}"
          f" · 계통 {len(parts.get('system') or ())}"
          f" · 기계실 {len(parts.get('machineroom') or ())}"
          f"  /  결합 절점 {len(nodes)} · 배관 {len(pipes)}")

    # ── E1: 어디에도 없는 라벨
    unc = [str(n.get("label")) for n in nodes
           if str(n.get("label")) not in of]
    print(f"  [E1] ★unclassified {len(unc)} {unc[:8]}")
    for lab in unc[:4]:
        n = next(x for x in nodes if str(x.get("label")) == lab)
        print(f"        {lab}: x={n.get('x')} y={n.get('y')}"
              f" elev={n.get('elevation')}")
    # 거짓 seam — 미분류 노드에 붙은 배관이 이음매로 보이는가
    fake = [str(p.get("label")) for p in pipes
            if (of.get(str(p.get("in")), "plan")
                != of.get(str(p.get("out")), "plan"))
            and (str(p.get("in")) in unc or str(p.get("out")) in unc)]
    print(f"  [E1] ★미분류 탓의 거짓 seam {len(fake)} {fake[:6]}")

    # ── E2: 배치 폴백 기록
    print(f"  [E2] layout_status = {got.get('layout_status')}")

    # ── E3: 부위별 bbox · 이음매 거리
    print("  [E3] 부위별 bbox")
    base = None
    for kind in ("plan", "system", "machineroom"):
        sub = [n for n in nodes if of.get(str(n.get("label"))) == kind]
        bb = _bbox(sub)
        if bb is None:
            print(f"        {kind:12} —")
            continue
        if kind == "plan":
            base = bb[4] or 1.0
        print(f"        {kind:12} span {bb[4]:10.0f}"
              f"  (평면의 {bb[4] / (base or 1):.2f}배)"
              f"  x[{bb[0]:.0f},{bb[1]:.0f}] y[{bb[2]:.0f},{bb[3]:.0f}]")
    whole = _bbox(nodes)
    print(f"        {'전체':12} span {whole[4]:10.0f}"
          f"  (평면의 {whole[4] / (base or 1):.2f}배)")

    ck = got.get("checks") or {}
    for k in ("anchor_gap", "pump_seam", "bbox_span_mm",
              "bbox_ratio_to_plan", "unclassified_n", "part_bbox"):
        print(f"  [E3] {k} = {ck.get(k)}")

    # 이음매 — 굽기 전 좌표
    at = {str(n.get("label")): (float(n.get("x", 0) or 0),
                                float(n.get("y", 0) or 0)) for n in nodes}
    elev = {str(n.get("label")): float(n.get("elevation", 0) or 0)
            for n in nodes}
    pj = got.get("pump_junction")
    if pj and str(pj) in at:
        mrset = set(parts.get("machineroom") or ())
        for p in pipes:
            a, b = str(p.get("in")), str(p.get("out"))
            if pj not in (a, b):
                continue
            other = b if a == pj else a
            if other not in mrset or other not in at:
                continue
            d = math.dist(at[pj], at[other])
            print(f"  [E3] pump_seam {pj}–{other} 좌표 {d:.1f} mm"
                  f" · 표 {p.get('length')} m"
                  f" · 비 {d / 1000.0 / max(1e-9, float(p.get('length') or 0)):.2f}")

    # ── 굽은 뒤 — 이음매 두 점이 같은가
    from routes.module_f.merge import bake_combined_iso
    rep: dict = {}
    iso, _edges = bake_combined_iso(got, report=rep)
    print(f"  [E3] seam_check = "
          f"{ {k: f'{v:.3e}' for k, v in (rep.get('seam_check') or {}).items()} }"
          f"  (기대 <1e-6)")
    isoat = {str(n.get("label")): (float(n.get("x", 0) or 0),
                                   float(n.get("y", 0) or 0)) for n in iso}
    seam_pipes = [p for p in pipes
                  if of.get(str(p.get("in")), "plan")
                  != of.get(str(p.get("out")), "plan")]
    print(f"  [E3] 굽은 뒤 절점 {len(iso)} · 부위가 갈리는 배관 {len(seam_pipes)}")
    for p in seam_pipes[:6]:
        a, b = str(p.get("in")), str(p.get("out"))
        if a in isoat and b in isoat:
            print(f"        {p.get('label')} {a}({of.get(a, '—')})"
                  f"–{b}({of.get(b, '—')})"
                  f" · 굽은 뒤 {math.dist(isoat[a], isoat[b]):.1f}"
                  f" · 굽기 전 {math.dist(at[a], at[b]):.1f}")

    # ── 총괄: 좌표 ÷ 1000 이 표 length 와 5% 넘게 다른 배관
    bad = []
    for p in pipes:
        a, b = str(p.get("in")), str(p.get("out"))
        ln = float(p.get("length") or 0)
        if a not in at or b not in at or ln <= 0.01:
            continue
        d = math.dist(at[a], at[b]) / 1000.0
        if abs(d - ln) > max(0.05 * ln, 0.05):
            # ★평면 좌표는 2D 다. 표고차만 있는 접속관(헤드 드롭)은 좌표
            #   거리가 0 이어도 표 길이가 있는 것이 **정상**이다 — 갈라 센다.
            dz = abs(elev.get(a, 0.0) - elev.get(b, 0.0))
            bad.append((str(p.get("label")), round(ln, 3), round(d, 3), dz))
    # ★어느 부위에서 나는지까지 갈라 본다. 계통도는 «schematic 수직 막대» 로
    #   일부러 균등 배치하므로 좌표가 표 길이와 다른 것이 정상이다 — 그 정상을
    #   평면·기계실의 진짜 어긋남과 섞어 세면 숫자가 아무 말도 못 한다.
    where: dict = {}
    for lab, _ln, _d, _dz in bad:
        p2 = next(q for q in pipes if str(q.get("label")) == lab)
        k = of.get(str(p2.get("in")), "?")
        k2 = of.get(str(p2.get("out")), "?")
        where[k if k == k2 else f"{k}|{k2}"] = where.get(
            k if k == k2 else f"{k}|{k2}", 0) + 1
    print(f"  [총괄] 좌표↔length 5% 초과 배관 {len(bad)} / {len(pipes)}"
          f"  · 부위별 {where}")
    dz_only = [r for r in bad if r[3] > 0.01 and abs(r[3] - r[1]) <= 0.05]
    print(f"         그중 «표고차뿐인 접속관»(정상) {len(dz_only)}"
          f" · 나머지 {len(bad) - len(dz_only)}")
    for row in bad[:6]:
        print(f"        {row[0]}: 표 {row[1]} m · 좌표 {row[2]} m"
              f" · 표고차 {row[3]:.3f} m")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

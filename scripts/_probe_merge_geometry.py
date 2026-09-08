# -*- coding: utf-8 -*-
"""[결합망 ②] 그려진 결합망의 «길이·수직·이음매» 를 잰다.

사용자 지적: 「계통도 쪽이랑 상하향식 길이가 많이 쪼개져서 깨져 있어.
공통 노드를 연결하는 과정에서 누수가 난 것 같아. 계통도도 수직으로
표현되어야 하는데 기울어져 있고.」

의심 셋을 각각 숫자로 가른다:

    ① 계통도(라이저) — 그려진 구간 길이가 표의 길이에 **비례**하는가.
       (`_layout_riser_as_schematic` 은 균등 간격이다 — 4 m 층과 0.3 m
        밸브 구간이 같은 길이로 그려지면 «쪼개져 깨진» 그림이 된다.)
    ② 수직 — 평면 보기에서 라이저가 한 x 에 서 있는가 · 아이소에서도
       수직으로 남는가(회전을 그대로 먹이면 기울어진다).
    ③ 이음매 — 기준점(10)·펌프 접속의 두 좌표계가 실제로 한 점에서
       만나는가(떨어져 있으면 «누수»).

    python scripts/_probe_merge_geometry.py
"""
from __future__ import annotations

import math
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def _load(name):
    import importlib.util
    path = ROOT / "tests" / "test_module_f_merge.py"
    spec = importlib.util.spec_from_file_location("_fx", str(path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return getattr(mod, name)


def _riser_real(n: int = 8) -> dict:
    """길이가 **서로 다른** 입상관 — 균등이면 비례가 깨지는 것이 보이게."""
    labels = ["1"] + [f"n{i}" for i in range(2, n)] + ["10"]
    lens = [4.0, 0.3, 3.5, 0.5, 4.0, 0.2, 2.5][: n - 1]
    nodes = [{"label": lab, "x": 0, "y": i * 1000, "elevation": float(i)}
             for i, lab in enumerate(labels)]
    nodes[0]["io_node"] = "Input"
    pipes = [{"label": f"r{i}", "in": labels[i], "out": labels[i + 1],
              "dia": 100, "length": lens[i]} for i in range(n - 1)]
    return {"nodes": nodes, "pipes": pipes, "av_node_label": "10"}


def main() -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    from routes.module_f.merge import merge_network

    got = merge_network(_load("_sample")(), riser=_riser_real(),
                        mode="lsp_gravity")
    c = got["combined"]
    at = {str(n["label"]): (float(n["x"]), float(n["y"]))
          for n in c.nodes}
    parts = got["parts"]
    sysset = set(parts["system"])

    print("■ ① 라이저 — 그려진 길이 ∝ 표 길이인가")
    rows = []
    for p in c.pipes:
        a, b = str(p.get("in")), str(p.get("out"))
        if a in sysset or b in sysset:
            if a in at and b in at:
                drawn = math.dist(at[a], at[b])
                rows.append((str(p.get("label")), float(p.get("length") or 0),
                             drawn))
    tot_len = sum(r[1] for r in rows) or 1.0
    tot_drawn = sum(r[2] for r in rows) or 1.0
    bad = 0
    for lab, ln, dr in rows:
        want = ln / tot_len
        gotr = dr / tot_drawn
        mark = ""
        if abs(want - gotr) > 0.02:
            bad += 1
            mark = "  ★비례 깨짐"
        print(f"    {lab:>4} 표 {ln:5.1f} m ({want * 100:5.1f}%)"
              f" · 그림 {dr:8.1f} ({gotr * 100:5.1f}%){mark}")
    print(f"    비례가 2%p 넘게 깨진 구간 {bad} / {len(rows)}")

    print("■ ② 수직 — 라이저 x 가 한 값인가 (평면)")
    xs = sorted({at[lab][0] for lab in sysset if lab in at})
    print(f"    라이저 x 값 종류 {len(xs)}: {xs[:5]}")

    print("■ ③ 이음매 — 두 좌표계가 한 점에서 만나는가")
    # 기준점 10: 라이저 AV 와 평면도 소스가 같은 라벨로 병합됐다.
    # 이음매 배관(라이저 마지막 → 10, 평면 10 → 11)의 그려진 길이를 본다.
    for p in c.pipes:
        a, b = str(p.get("in")), str(p.get("out"))
        ka = "sys" if a in sysset else "plan"
        kb = "sys" if b in sysset else "plan"
        if ka != kb and a in at and b in at:
            print(f"    {p.get('label')}: {a}({ka}) → {b}({kb})"
                  f" · 표 {p.get('length')} m"
                  f" · 그림 {math.dist(at[a], at[b]):.1f}")

    # ── 아이소 — «화면 라우트 그대로» 굽는다 (미리보기 라우트를 직접 부른다)
    print("■ ② -b 아이소 — 라이저가 수직으로 남는가 (미리보기 라우트)")
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as cl:
        with cl.session_transaction() as ss:
            ss["authed"] = True
        from routes.module_f.jobs import _new_session
        sess = _new_session()
        sess["merged"] = got
        d = cl.get(f"/api/module-f/merge/preview?sid={sess['id']}"
                   f"&iso=1").get_json()
        v = d["view"]
        at2 = {n["label"]: (n["x"], n["y"]) for n in v["nodes"]}
        parts2 = {n["label"]: n["part"] for n in v["nodes"]}
        xs2 = sorted({round(at2[lab][0], 1) for lab in at2
                      if parts2.get(lab) == "system"})
        print(f"    아이소 뒤 라이저 x 값 종류 {len(xs2)}"
              f" — {'수직 유지' if len(xs2) == 1 else '★기울어짐'}")
        if len(xs2) > 1:
            print(f"    x 퍼짐 {xs2[-1] - xs2[0]:.1f} (막대가 사선이 됐다)")
        # 이음매 — 아이소에서도 라이저↔평면이 붙어 있는가(기준점 좌표 하나).
        seam = [p for p in v["pipes"] if p["part"] == "seam"]
        for p in seam:
            dd = math.dist(at2[p["a"]], at2[p["b"]])
            print(f"    이음매 {p['a']}→{p['b']} 그림 거리 {dd:.1f}")
        # 라이저 비례가 아이소에서도 유지되는가.
        rows2 = []
        for p in v["pipes"]:
            if parts2.get(p["a"]) == "system" and parts2.get(p["b"]) == "system":
                rows2.append((float(p.get("len_m") or 0),
                              math.dist(at2[p["a"]], at2[p["b"]])))
        t1 = sum(r[0] for r in rows2) or 1.0
        t2 = sum(r[1] for r in rows2) or 1.0
        bad2 = sum(1 for ln, dr in rows2
                   if abs(ln / t1 - dr / t2) > 0.02)
        print(f"    아이소 라이저 비례 깨짐 {bad2} / {len(rows2)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

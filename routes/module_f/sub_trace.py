# -*- coding: utf-8 -*-
"""[오너 2026-09-22] 계통도 칸 — 레이어 세 묶음과 경로 추적 레이어(★).

B1F 처럼 기계실과 평면도가 같은 층인 도면은 계통도 칸에도 평면도를 올려 그
안에서 펌프 → 알람밸브 경로를 뽑는다. 그런데 계통도 칸은 CAD 에서 꺼 둔
레이어까지 읽는다 — 꺼 둔 레이어에 배관이 그려진 계통도가 있어서다(MF-004 의
`SP-L`). 그래서 평면도를 올리면 헤드 반경 원 · 소화전 반경 · 문자 · 구역선이
한 화면에 섞이고, 이름 사전이 고른 «배관» 20개 가운데 진짜 배관은 4개뿐이라
경로가 구역선을 타고 꼬였다(B1F 실측 평면 521 m).

여기서 하는 일은 셋이다.

  ① 레이어를 세 묶음으로 나눈다 — 배관망 · 건축 · 숨김(`layer_rows`).
     화면 기본값은 앞의 둘만 보인다. 숨김도 목록에 남아 켜면 보인다 —
     **화면 표시만** 바꾼다(추적과 무관).
  ② 경로 추적 레이어(★)를 정한다(`plan_trace`). 자동 레이어에 숨김·헤드
     레이어가 섞인 도면만 «켜진 배관 레이어» 로 좁힌다. 좁힌 망이 멀리 끊겨
     억지 연결이 하나라도 필요하면 — 배관이 다른 레이어에도 그려져 있다는
     뜻이다(MF-004 실측 254개) — 지금 방식 그대로 둔다. 섞인 것이 없는
     도면(대명동 계통도)도 지금 그대로다.
  ③ 좁힌 배관의 T 접속을 잇는다(`split_t_junctions`). 한 배관의 **끝**이 다른
     배관 **한가운데**에 닿는 자리다. A 의 그래프는 끝점끼리만 잇기 때문에 이
     자리가 끊기고, 경로가 옆 가지관으로 돌아갔다(B1F: 주배관 끝이 세로 1차
     주배관 한가운데에 닿는 자리 — 오너가 빨간 선으로 짚은 곳).

★좁히기(②)와 T 접속(③)은 **함께만** 켜진다. 지금 방식이 잘 되는 도면의
  경로는 한 글자도 바꾸지 않는다 — 그 도면은 ② 에서 «그대로» 로 떨어진다.
"""
from __future__ import annotations

import math
from collections import defaultdict

# 이름에 이것이 있으면 배관망 묶음이다 — 밸브·부속·입상·알람·펌프 기호 레이어.
NET_KEYWORDS = ("밸브", "VALVE", "V/V", "FLX", "FLEX", "후렉", "부속", "FITTING",
                "입상", "RISER", "알람", "ALARM", "펌프", "PUMP", "감압")
# 이름에 이것이 있으면 배관이 아니라 «범위» 표시다 — 헤드 반경 · 구역선.
NOT_NET = ("반경", "RADIUS", "범위", "ZONE", "존", "구역")

GROUP_ORDER = ("net", "arch", "etc")
GROUP_LABELS = {"net": "배관망", "arch": "건축", "etc": "숨긴 레이어"}

# T 접속 허용치 = A 그래프의 «같은 점» 자(`snap_eps`)의 5 %.
#   B1F(축척비 0.743): 37.2 mm × 0.05 = 1.9 mm. CAD 에서 끝점을 선 위에 붙여
#   그리면 거리가 0 이다(실측: 주배관 끝 ↔ 세로 주배관 0.0 mm). 가지관의 짧은
#   눈금선 끝은 주배관에서 31 mm 떨어져 있다 — 허용치를 snap_eps(37 mm) 그대로
#   두면 그 눈금을 T 로 읽어 주배관이 31 mm 씩 꺾였다(실측으로 걸러 냈다).
T_JOINT_FRACTION = 0.05


def layer_group(name, *, visible, cat_name, cat_geo):
    """레이어 하나 → (묶음, 까닭). 묶음은 `net` · `arch` · `etc`."""
    u = str(name).upper()
    cats = {cat_name, cat_geo}
    if not visible:
        return "etc", "CAD 꺼 둠"
    if "EXCLUDE" in cats or "TEXT" in cats:
        return "etc", "문자·제외"
    if any(k in u for k in NOT_NET):
        return "etc", "구역·반경"
    if cats & {"PIPE", "HEAD", "ALARM"}:
        return "net", "배관·헤드"
    if any(k.upper() in u for k in NET_KEYWORDS):
        return "net", "밸브·부속"
    if "ARCH" in cats:
        return "arch", "건축"
    return "etc", "분류 안 됨"


def layer_rows(entities, parsed) -> list[dict]:
    """계통도 칸의 레이어 표 — 묶음 · 까닭 · 도형 수 · CAD 켜짐 · 자동 추적 여부.

    `cat` 은 기하 교정을 거친 분류다(`categorize_layers` — 이름은 배관인데 내용이
    닫힌 기호뿐인 레이어를 내린다). 묶음은 이름 분류와 기하 분류를 **함께** 본다.
    """
    from remote30_prototype import (_auto_pipe_layer_filter, _categorize_layer,
                                    categorize_layers)
    geo = categorize_layers(entities)
    auto = _auto_pipe_layer_filter(entities)
    n_by: dict = defaultdict(int)
    for en in (entities or ()):
        n_by[str(en.get("l") or "0")] += 1
    vis: dict = {}
    for ly in ((parsed or {}).get("layers") or ()):
        vis[str(ly.get("name"))] = bool(ly.get("visible", True))
    rows = []
    for name in sorted(set(n_by) | set(vis)):
        if n_by.get(name, 0) == 0:
            continue                   # 도형이 없는 레이어는 표에 둘 까닭이 없다
        cn = _categorize_layer(name)
        cg = geo.get(name) or cn
        group, why = layer_group(name, visible=vis.get(name, True),
                                 cat_name=cn, cat_geo=cg)
        rows.append({"layer": name, "group": group, "why": why,
                     "n": int(n_by[name]), "visible": vis.get(name, True),
                     "cat": cg, "cat_name": cn, "auto": name in auto})
    rows.sort(key=lambda r: (GROUP_ORDER.index(r["group"]), -r["n"], r["layer"]))
    return rows


def _segments(entities, layers):
    """(도형 번호, 마디 번호, (x1, y1, x2, y2)) — `layers` 의 선·폴리선 마디."""
    out = []
    for i, en in enumerate(entities or ()):
        if en.get("l") not in layers:
            continue
        t = en.get("t")
        if t == "L":
            p = en.get("p") or ()
            if len(p) >= 4:
                out.append((i, 0, (float(p[0]), float(p[1]),
                                   float(p[2]), float(p[3]))))
        elif t == "PL":
            pts = en.get("p") or ()
            for k in range(len(pts) - 1):
                a, b = pts[k], pts[k + 1]
                if len(a) >= 2 and len(b) >= 2:
                    out.append((i, k, (float(a[0]), float(a[1]),
                                       float(b[0]), float(b[1]))))
    return out


def split_t_junctions(entities, layers, tol):
    """`layers` 의 배관 마디를 «다른 배관의 끝이 닿는 한가운데 자리» 에서 자른다.

    반환 `(새 도형 목록, 통계)`. 자른 도형만 선(`L`) 조각으로 바뀌고 나머지
    도형은 **같은 객체** 그대로다(문자 · 호 · 다른 레이어 전부).

    ★끝이 닿은 자리의 좌표를 **그 끝점 좌표 그대로** 쓴다(선 위로 투영하지
      않는다). 그래야 그래프에서 두 배관이 정확히 한 점을 나눈다. 허용치(`tol`)가
      작아(B1F 1.9 mm) 선이 옮겨지는 폭도 그 안이다.
    ★엇갈려 지나가는 두 배관(어느 쪽도 끝나지 않는 X 교차)은 **자르지 않는다** —
      호 기호 없는 교차는 잇지 않는다는 평면도 규칙과 같다.
    """
    layers = set(layers or ())
    segs = _segments(entities, layers)
    if not segs or tol <= 0:
        return list(entities or ()), {"split_entities": 0, "cuts": 0}
    cell = max(float(tol) * 20.0, 1000.0)
    grid: dict = defaultdict(list)
    for si, (_i, _k, (x1, y1, x2, y2)) in enumerate(segs):
        cx0, cx1 = math.floor(min(x1, x2) / cell), math.floor(max(x1, x2) / cell)
        cy0, cy1 = math.floor(min(y1, y2) / cell), math.floor(max(y1, y2) / cell)
        if (cx1 - cx0 + 1) * (cy1 - cy0 + 1) <= 4096:
            cells = [(cx, cy) for cx in range(cx0, cx1 + 1)
                     for cy in range(cy0, cy1 + 1)]
        else:                          # 아주 긴 사선 — 선을 따라 칸을 걷는다
            n = int(math.hypot(x2 - x1, y2 - y1) / (cell / 2.0)) + 1
            cells = {(math.floor((x1 + (x2 - x1) * j / n) / cell),
                      math.floor((y1 + (y2 - y1) * j / n) / cell))
                     for j in range(n + 1)}
        for c in cells:
            grid[c].append(si)
    cuts: dict = defaultdict(set)      # 마디 → {(t, x, y)}
    for si, (_i, _k, (x1, y1, x2, y2)) in enumerate(segs):
        for (px, py) in ((x1, y1), (x2, y2)):
            c = (math.floor(px / cell), math.floor(py / cell))
            for sj in grid.get(c, ()):
                if sj == si:
                    continue
                ax, ay, bx, by = segs[sj][2]
                dx, dy = bx - ax, by - ay
                l2 = dx * dx + dy * dy
                if l2 <= 0.0:
                    continue
                ln = math.sqrt(l2)
                t = ((px - ax) * dx + (py - ay) * dy) / l2
                if t * ln <= tol or (1.0 - t) * ln <= tol:
                    continue           # 끝 가까이 — 끝끼리 잇는 보통 이음이다
                if abs((px - ax) * dy - (py - ay) * dx) / ln > tol:
                    continue
                cuts[sj].add((t, px, py))
    if not cuts:
        return list(entities), {"split_entities": 0, "cuts": 0}
    by_ent: dict = defaultdict(list)
    for si, (i, k, g) in enumerate(segs):
        by_ent[i].append((k, si, g))
    hit = {segs[si][0] for si in cuts}
    out, n_cuts = [], 0
    for i, en in enumerate(entities):
        if i not in hit:
            out.append(en)
            continue
        for _k, si, (x1, y1, x2, y2) in sorted(by_ent[i]):
            pts = [(0.0, x1, y1)] + sorted(cuts.get(si, ())) + [(1.0, x2, y2)]
            n_cuts += len(pts) - 2
            for (ta, xa, ya), (tb, xb, yb) in zip(pts, pts[1:]):
                if math.hypot(xb - xa, yb - ya) <= 1e-6:
                    continue
                out.append({"t": "L", "l": en.get("l"), "p": [xa, ya, xb, yb]})
    return out, {"split_entities": len(hit), "cuts": n_cuts}


def plan_trace(entities, rows) -> dict:
    """경로 추적 레이어(★)를 정한다 — 올릴 때 한 번.

    반환::

        {"mode": "pipe" | "auto",
         "layers": [★ 레이어…] | None,        # None = 지금 방식(자동 전체)
         "reason": 사람에게 보일 한 줄,
         "junk": [자동에 섞인 숨김·헤드 레이어…],
         "candidates": [켜진 배관 레이어…],
         "forced": 좁힌 망의 억지 연결 수 | None,
         "split": {"split_entities", "cuts"} | None,
         # "pipe" 일 때만 — 서버 안에서만 쓰고 화면에는 안 보낸다
         "entities": T 접속을 자른 도형 목록,
         "graph": (graph, edge_len, stats)}
    """
    from remote30_prototype import (SNAP_TOL_MM, _drawing_scale_ratio,
                                    build_system_graph)
    by = {r["layer"]: r for r in rows}
    auto = sorted(r["layer"] for r in rows if r.get("auto"))
    junk = sorted(nm for nm in auto
                  if by[nm]["group"] == "etc"
                  or "HEAD" in (by[nm]["cat"], by[nm]["cat_name"]))
    out = {"mode": "auto", "layers": None, "junk": junk, "candidates": [],
           "forced": None, "split": None}
    if not junk:
        out["reason"] = "자동으로 고른 배관 레이어에 숨김·헤드 레이어가 없어 지금 방식 그대로 추적합니다"
        return out
    cand = sorted(nm for nm in auto
                  if by[nm]["visible"] and by[nm]["group"] == "net"
                  and "PIPE" in (by[nm]["cat"], by[nm]["cat_name"]))
    out["candidates"] = cand
    if not cand:
        out["reason"] = "자동 레이어 중 켜진 배관 레이어가 없어 지금 방식 그대로 추적합니다"
        return out
    lines = [en for en in entities
             if en.get("t") in ("L", "PL") and en.get("l") in set(cand)]
    snap_eps = SNAP_TOL_MM * _drawing_scale_ratio(lines)
    ents2, st = split_t_junctions(entities, cand, T_JOINT_FRACTION * snap_eps)
    graph, edge_len, stats = build_system_graph(ents2, layer_filter=set(cand),
                                                force_connect=True)
    out["split"] = st
    out["forced"] = int(stats.get("forced_bridges") or 0)
    if out["forced"]:
        out["reason"] = (f"배관 레이어만으로는 멀리 끊긴 곳을 억지로 이어야 해서"
                         f"({out['forced']}곳) 지금 방식 그대로 추적합니다")
        return out
    out.update({"mode": "pipe", "layers": cand, "entities": ents2,
                "graph": (graph, edge_len, stats),
                "reason": (f"자동 레이어 {len(auto)}개 중 켜진 배관 {len(cand)}개로만 "
                           f"추적합니다 — 헤드·밸브 기호·구역선·문자는 쓰지 않습니다")})
    return out


def public_trace(tr) -> dict | None:
    """화면에 보낼 몫 — 도형·그래프는 빼고 말만."""
    if not tr:
        return None
    return {k: tr.get(k) for k in ("mode", "layers", "reason", "junk",
                                   "candidates", "forced", "split")}


def public_rows(rows) -> list[dict]:
    """화면에 보낼 레이어 표 — 내부 칸(`cat_name`)은 뺀다."""
    return [{k: r[k] for k in ("layer", "group", "why", "n", "visible", "cat",
                               "auto")} for r in (rows or ())]

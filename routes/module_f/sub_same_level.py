# -*- coding: utf-8 -*-
"""[오너 2026-09-22] 계통도 칸 «같은 층» 추출의 뒤처리 — 평면도로 그린 계통도.

B1F 처럼 기계실과 평면도가 같은 층이면 계통도 칸에도 평면도를 올린다. 그 도면의
위아래(Y)는 높이가 아니라 평면 방향이라 표고 기준을 «같은 층» 으로 고른다. 이
모듈은 그 선택의 뒤처리만 맡는다 — «도면» 기준(세로로 그린 계통도)은 여기를
지나지 않으므로 한 글자도 안 바뀐다.

  ① 실측 길이(`measured_lengths`) — 추출은 한 직선 위 조각들을 배관 하나로 합치고
     (B1F r1 = 186.3 m), 층 경계 글자 자리에서 곧은 배관을 쪼갠다. 그런 배관은
     «조각 길이표» 에 없어서 추출 전체가 «r1: 표고 보정 전 실측 길이를 찾을 수
     없습니다» 로 멈췄다. 그 직선 구간을 그래프 조각이 **빈틈없이 덮으면** 그
     길이를 쓴다. 틈이 있으면 지금처럼 멈춘다 — 길이를 지어내지 않는다.
  ② 층 경계 조각 다시 잇기 — 평면도의 입상 표기 글자(«지하1층<-지상1층» 등)를
     층 이름으로 읽어 곧은 배관을 잘게 쪼갠 자리를 한 배관으로 되잇는다(같은
     관경일 때만). 같은 층에서는 층이 없으므로 그 마디가 아무것도 뜻하지 않는다.
  ③ 알람밸브 꼬리 — 알람밸브 조립 상세(입면으로 그린 이중선)를 따라 오르내린
     꼬리를, 알람밸브 둘레에 들어온 첫 노드에서 알람밸브 노드로 곧장 잇는다
     (B1F: 이중선 2.1 m 오르내림 → 0.62 m). 곧은 거리의 1.5 배를 넘을 때만.
  ④ 호 기호 우회 — 경로 배관 위에 원호 두 개가 마주 보면(덱 「실제길이 A」) 그
     사이를 0.5 m 올린다: 양 끝에 세로관 0.5 m, 꺾이는 자리 넷은 엘보 4개.
     평면도 규칙(`convert/arc_jog.py`)과 같은 자 — 열린 쪽 판정은 `is_open`,
     높이는 `dto.BRANCH_DEFAULT_M`. 짝 없는 호는 손대지 않고 센다.
  ⑤ 부속 — 꺾인 자리 엘보(45°·90°), 티(도면에서 셋 이상 이어진 자리)에서 꺾이면
     분류티. 판정은 평면도와 같은 함수 한 곳(`design.fitting.build_fittings`)이
     한다. 직류티는 계상하지 않는다(`CALCULATION_FITTINGS`).
"""
from __future__ import annotations

import math
from collections import defaultdict

STRAIGHT_DEG = 2.0        # A 의 `_collapse_collinear_nodes(angle_tol_deg=2.0)` 와 같은 자
AV_TAIL_R_MM = 1500.0     # 알람밸브 둘레 — 도면 축척비를 곱해 쓴다(B1F 1.1 m: 조립 상세 0.75 m 를 덮는다)
AV_TAIL_RATIO = 1.5       # 꼬리가 곧은 거리의 이 배를 넘으면 «오르내린 꼬리»
HOP_MAX_RUN_M = 60.0      # `arc_jog.MAX_RUN_M` 과 같다
HOP_MAX_STEPS = 64        # `arc_jog.MAX_STEPS` 와 같다
HOP_R_MIN_MM = 100.0      # 이보다 작은 호는 끊김 표시(파단선 r 52)·밸브 기호의 일부다 —
                          #   덱의 우회·괄호 호는 r ≈ 175 mm (B1F 실측)
HOP_R_MAX_MM = 600.0
HOP_SWEEP = (45.0, 315.0)  # 온 원(360°)과 점 같은 호는 기호가 아니다


def _key(a, b):
    return (min(a, b), max(a, b))


# ── ① 실측 길이 ────────────────────────────────────────────────────────────
def measured_lengths(result, edge_len, stats) -> dict:
    """«같은 층» 이 쓸 실측 길이표 {(끝점, 끝점): mm}.

    조각 길이표를 그대로 싣고(종전), 표에 없는 경로 배관만 더한다 — 그 직선
    구간을 조각(다리 포함 — «도면» 기준도 같은 값을 쓴다)이 틈 없이 덮을 때만.
    """
    measured: dict = {}
    for (a, b), length in (edge_len or {}).items():
        pa = (round(a[0]), round(a[1]))
        pb = (round(b[0]), round(b[1]))
        measured[_key(pa, pb)] = length
    nodes = {str(n["label"]): n for n in (result.get("nodes") or ())}
    want = []
    for p in (result.get("pipes") or ()):
        a, b = nodes.get(str(p.get("in"))), nodes.get(str(p.get("out")))
        if a is None or b is None or p.get("length_measured_mm") is not None:
            continue
        pa = (float(a["x"]), float(a["y"]))
        pb = (float(b["x"]), float(b["y"]))
        if _key(pa, pb) not in measured:
            want.append((pa, pb))
    if not want:
        return measured
    tol = max(float((stats or {}).get("snap_eps_mm") or 0.0), 1.0)
    cell = 2000.0
    grid: dict = defaultdict(list)
    edges = list((edge_len or {}).keys())
    for ei, (a, b) in enumerate(edges):
        for cx in range(math.floor(min(a[0], b[0]) / cell),
                        math.floor(max(a[0], b[0]) / cell) + 1):
            for cy in range(math.floor(min(a[1], b[1]) / cell),
                            math.floor(max(a[1], b[1]) / cell) + 1):
                grid[(cx, cy)].append(ei)
    for pa, pb in want:
        ln = math.dist(pa, pb)
        if ln <= 0.0:
            continue
        ux, uy = (pb[0] - pa[0]) / ln, (pb[1] - pa[1]) / ln
        seen, spans = set(), []
        for cx in range(math.floor((min(pa[0], pb[0]) - tol) / cell),
                        math.floor((max(pa[0], pb[0]) + tol) / cell) + 1):
            for cy in range(math.floor((min(pa[1], pb[1]) - tol) / cell),
                            math.floor((max(pa[1], pb[1]) + tol) / cell) + 1):
                for ei in grid.get((cx, cy), ()):
                    if ei in seen:
                        continue
                    seen.add(ei)
                    ts = []
                    for (x, y) in edges[ei]:
                        if abs(-(x - pa[0]) * uy + (y - pa[1]) * ux) > tol:
                            break
                        ts.append((x - pa[0]) * ux + (y - pa[1]) * uy)
                    else:
                        lo, hi = max(0.0, min(ts)), min(ln, max(ts))
                        if hi > lo:
                            spans.append((lo, hi))
        spans.sort()
        covered, cur = 0.0, None
        for lo, hi in spans:
            if cur is None or lo > cur[1]:
                if cur is not None:
                    covered += cur[1] - cur[0]
                cur = [lo, hi]
            else:
                cur[1] = max(cur[1], hi)
        if cur is not None:
            covered += cur[1] - cur[0]
        if ln - covered <= tol:
            measured[_key(pa, pb)] = ln        # 틈 없이 덮였다 — 그 직선 길이
    return measured


# ── 경로 한 줄로 세우기 ─────────────────────────────────────────────────────
def _chain(result):
    """급수원 → 알람밸브 순서의 (노드 라벨 열, 배관 열). 한 줄이 아니면 None."""
    nodes = {str(n["label"]): n for n in (result.get("nodes") or ())}
    pipes = list(result.get("pipes") or ())
    start = str(result.get("input_node_label") or "")
    if start not in nodes:
        start = next((str(n["label"]) for n in nodes.values()
                      if str(n.get("io_node", "")).lower() == "input"), "")
    end = str(result.get("av_node_label") or "")
    if start not in nodes or end not in nodes:
        return None
    by_node: dict = defaultdict(list)
    for p in pipes:
        by_node[str(p["in"])].append(p)
        by_node[str(p["out"])].append(p)
    order_n, order_p, used = [start], [], set()
    cur = start
    while cur != end:
        nxt = [p for p in by_node.get(cur, ()) if id(p) not in used]
        if len(nxt) != 1:
            return None
        p = nxt[0]
        used.add(id(p))
        cur = str(p["out"]) if str(p["in"]) == cur else str(p["in"])
        if str(p["in"]) != order_n[-1]:           # 방향을 급수원 → 알람밸브로
            p["in"], p["out"] = p["out"], p["in"]
            p["elev"] = -float(p.get("elev") or 0.0)
        order_n.append(cur)
        order_p.append(p)
        if len(order_n) > len(nodes) + 1:
            return None
    if len(order_p) != len(pipes):
        return None
    return order_n, order_p


def _xy(nodes, lab):
    n = nodes[lab]
    return float(n["x"]), float(n["y"])


def _turn_deg(a, m, b):
    ux, uy = m[0] - a[0], m[1] - a[1]
    vx, vy = b[0] - m[0], b[1] - m[1]
    lu, lv = math.hypot(ux, uy), math.hypot(vx, vy)
    if lu <= 1e-9 or lv <= 1e-9:
        return 0.0
    c = max(-1.0, min(1.0, (ux * vx + uy * vy) / (lu * lv)))
    return math.degrees(math.acos(c))


def _same_run(p, q) -> bool:
    return (str(p.get("dia")) == str(q.get("dia"))
            and str(p.get("type")) == str(q.get("type"))
            and str(p.get("c")) == str(q.get("c")))


# ── ② 곧은 조각 다시 잇기 ────────────────────────────────────────────────────
def _merge_straight(result, order_n, order_p):
    nodes = {str(n["label"]): n for n in result["nodes"]}
    ends = {order_n[0], order_n[-1]}
    keep_n, keep_p, merged = [order_n[0]], [], 0
    for i, p in enumerate(order_p):
        m = order_n[i]
        if keep_p and m not in ends:
            q = keep_p[-1]
            a = _xy(nodes, keep_n[-2])
            b = _xy(nodes, order_n[i + 1])
            if _same_run(q, p) and _turn_deg(a, _xy(nodes, m), b) <= STRAIGHT_DEG:
                q["out"] = p["out"]
                q["length"] = round(float(q["length"]) + float(p["length"]), 3)
                if q.get("inferred_length") is not None and p.get("inferred_length") is not None:
                    q["inferred_length"] = round(float(q["inferred_length"])
                                                 + float(p["inferred_length"]), 3)
                keep_n[-1] = order_n[i + 1]
                merged += 1
                continue
        keep_p.append(p)
        keep_n.append(order_n[i + 1])
    gone = set(order_n) - set(keep_n)
    result["nodes"] = [n for n in result["nodes"] if str(n["label"]) not in gone]
    result["pipes"] = keep_p
    return keep_n, keep_p, merged


# ── ③ 알람밸브 꼬리 ─────────────────────────────────────────────────────────
def _trim_av_tail(result, order_n, order_p, scale):
    """알람밸브 둘레(반지름 1.5 m × 축척비) 안에서 오르내린 꼬리를 곧장 잇는다.

    둘레 안 노드 가운데 **알람밸브에 가장 가까운 쪽부터** 보아, 거기서 알람밸브까지
    남은 경로(배관 2개 이상)가 곧은 거리의 1.5 배를 넘는 첫 노드에서 자른다 —
    진짜 배관(입상관 등)은 남기고 오르내린 꼬리만 걷는다. 곧게 들어오거나 한 번
    꺾여 들어오는 경로(√2 ≈ 1.41 배)는 건드리지 않는다.
    """
    nodes = {str(n["label"]): n for n in result["nodes"]}
    av = _xy(nodes, order_n[-1])
    lim = AV_TAIL_R_MM * float(scale or 1.0)
    last = len(order_n) - 1
    k0 = last
    while k0 > 1 and math.dist(_xy(nodes, order_n[k0 - 1]), av) <= lim:
        k0 -= 1
    cut = None
    for k in range(last - 2, k0 - 1, -1):
        if k <= 0:
            break
        straight_mm = math.dist(_xy(nodes, order_n[k]), av)
        tail_m = sum(float(p["length"]) for p in order_p[k:])
        if straight_mm > 0.0 and tail_m * 1000.0 > AV_TAIL_RATIO * straight_mm:
            cut = (k, straight_mm, tail_m)
            break
    if cut is None:
        return order_n, order_p, None
    k, straight_mm, tail_m = cut
    tail = order_p[k:]
    base = order_p[k - 1]
    conn = dict(tail[0])
    for fld in ("dia", "type", "c", "dia_source", "dia_raw"):
        if fld in base:
            conn[fld] = base[fld]
    conn.update({"in": order_n[k], "out": order_n[-1],
                 "length": round(straight_mm / 1000.0, 3), "elev": 0.0,
                 "av_tail": True})
    conn.pop("inferred_length", None)
    info = {"from": order_n[k], "removed_pipes": len(tail),
            "removed_m": round(tail_m, 3), "straight_m": round(straight_mm / 1000.0, 3)}
    gone = set(order_n[k + 1:-1])
    result["nodes"] = [n for n in result["nodes"] if str(n["label"]) not in gone]
    new_p = order_p[:k] + [conn]
    result["pipes"] = new_p
    return order_n[:k + 1] + [order_n[-1]], new_p, info


# ── ④ 호 기호 우회 ─────────────────────────────────────────────────────────
def _arcs_on(entities, layers):
    out = []
    for en in (entities or ()):
        if en.get("t") != "A" or (layers and en.get("l") not in layers):
            continue
        c = en.get("c") or ()
        r = float(en.get("r") or 0.0)
        if len(c) < 2 or not (HOP_R_MIN_MM <= r <= HOP_R_MAX_MM):
            continue
        a = en.get("a") or (0.0, 360.0)
        sa = float(a[0])
        sweep = ((float(a[1]) if len(a) > 1 else sa + 360.0) - sa) % 360.0 or 360.0
        if not (HOP_SWEEP[0] <= sweep <= HOP_SWEEP[1]):
            continue
        out.append({"cx": float(c[0]), "cy": float(c[1]), "r": r,
                    "sa": sa, "sweep": sweep})
    return out


def _hops(result, order_n, order_p, arcs, phys, rise, tol):
    """경로 위 원호 짝 → 0.5 m 우회. 반환 (노드 열, 배관 열, 보고).

    ① 호마다 경로 배관 위 자리(급수원에서 잰 거리 s)와 «열린 쪽» 을 잰다 — 중심이
       배관 선 위(반지름의 30 % 안)에 있고, 배관을 따라 한쪽만 열린 호만 후보다.
    ② 급수원 쪽에서 «하류로 열린» 호 다음에 «상류로 열린» 호가 오면 짝이다
       (덱 그림 1 의 ⊂ … ⊃). 사이에 갈래(도면 3방향 이상)가 있으면 짝이 아니다.
    ③ 짝의 두 중심에서 배관을 자르고, 두 끝에 세로관(rise)을 세우고 사이를 올린다.
    """
    report = {"pairs": 0, "unpaired": 0, "rise_m": rise}
    if not arcs or rise <= 0.0:
        return order_n, order_p, report
    from services.cad_import.convert.main_walk import is_open
    nodes = {str(n["label"]): n for n in result["nodes"]}
    pts = [_xy(nodes, lab) for lab in order_n]
    cum = [0.0]
    for a, b in zip(pts, pts[1:]):
        cum.append(cum[-1] + math.dist(a, b))
    marks = []                            # (s, 향함 +1/-1, 배관 번호, 점)
    for arc in arcs:
        best = None
        for i, (a, b) in enumerate(zip(pts, pts[1:])):
            ln = math.dist(a, b)
            if ln <= 0.0:
                continue
            ux, uy = (b[0] - a[0]) / ln, (b[1] - a[1]) / ln
            t = (arc["cx"] - a[0]) * ux + (arc["cy"] - a[1]) * uy
            d = abs(-(arc["cx"] - a[0]) * uy + (arc["cy"] - a[1]) * ux)
            if t < -tol or t > ln + tol or d > max(tol, 0.3 * arc["r"]):
                continue
            if best is None or d < best[0]:
                best = (d, i, min(max(t, 0.0), ln), ux, uy)
        if best is None:
            continue
        _d, i, t, ux, uy = best
        fwd = math.degrees(math.atan2(uy, ux))
        o_f, o_b = is_open([arc], fwd), is_open([arc], fwd + 180.0)
        if o_f == o_b:
            continue                          # 양쪽 열림 · 양쪽 막힘 — 우회 기호가 아니다
        a = pts[i]
        marks.append((cum[i] + t, 1 if o_f else -1, i,
                      (a[0] + ux * t, a[1] + uy * t)))
    marks.sort()
    pairs, used = [], set()
    for j, (s1, f1, i1, _p1) in enumerate(marks):
        if j in used or f1 != 1:
            continue
        for k in range(j + 1, len(marks)):
            s2, f2, i2, _p2 = marks[k]
            if (s2 - s1) / 1000.0 > HOP_MAX_RUN_M or i2 - i1 > HOP_MAX_STEPS:
                break
            if f2 == 1:
                break                         # 다음 호가 같은 쪽을 보면 짝이 없다
            if s2 - s1 <= tol:
                break
            inner = [order_n[x] for x in range(i1 + 1, i2 + 1)
                     if cum[x] > s1 + tol and cum[x] < s2 - tol]
            if any(int(phys.get(lab, 2)) >= 3 for lab in inner):
                break                         # 사이에 갈래(티)가 있으면 우회가 아니다
            pairs.append((j, k))
            used.update((j, k))
            break
    report["pairs"] = len(pairs)
    report["unpaired"] += len(marks) - 2 * len(pairs)
    if not pairs:
        return order_n, order_p, report

    # 새 라벨 — 라이저의 n·r 번호를 잇는다(라벨 공간을 새로 만들지 않는다)
    nums = [int(lab[1:]) for lab in nodes if lab.startswith("n") and lab[1:].isdigit()]
    pnums = [int(str(p["label"])[1:]) for p in result["pipes"]
             if str(p["label"]).startswith("r") and str(p["label"])[1:].isdigit()]
    nx, px = [max(nums + [10])], [max(pnums + [0])]

    def new_node(xy, z, like):
        nx[0] += 1
        lab = f"n{nx[0]}"
        n = {k: v for k, v in nodes[like].items() if k not in ("io_node", "pressure_pa")}
        n.update({"label": lab, "x": int(round(xy[0])), "y": int(round(xy[1])),
                  "elevation": round(float(z), 3), "io_node": "No"})
        result["nodes"].append(n)
        nodes[lab] = n
        return lab

    def new_label():
        px[0] += 1
        return f"r{px[0]}"

    # ① 호 자리에서 배관을 자른다
    cuts = defaultdict(list)                 # 배관 번호 → [(배관 안 거리, 점, 꼬리표)]
    for idx, (j, k) in enumerate(pairs):
        for tag, (s, _f, i, xy) in ((("a", idx)), marks[j]), ((("b", idx)), marks[k]):
            cuts[i].append((s - cum[i], xy, tag))
    tag_at, new_n, new_p = {}, [order_n[0]], []
    for i, p in enumerate(order_p):
        a_lab, b_lab = order_n[i], order_n[i + 1]
        la = cum[i + 1] - cum[i]
        prev_lab, prev_t, first = a_lab, 0.0, True
        for t, xy, tag in sorted(cuts.get(i, ()), key=lambda c: c[0]):
            if t - prev_t <= tol:
                tag_at[tag] = prev_lab
                continue
            if la - t <= tol:
                tag_at[tag] = b_lab
                continue
            lab = new_node(xy, float(nodes[a_lab].get("elevation") or 0.0), a_lab)
            piece = dict(p)
            piece.update({"in": prev_lab, "out": lab,
                          "length": round(float(p["length"]) * (t - prev_t) / la, 3)})
            if not first:
                piece["label"] = new_label()
            piece.pop("inferred_length", None)
            new_p.append(piece)
            new_n.append(lab)
            tag_at[tag] = lab
            prev_lab, prev_t, first = lab, t, False
        last = dict(p)
        last.update({"in": prev_lab, "out": b_lab})
        if not first:
            last.update({"label": new_label(),
                         "length": round(float(p["length"]) * (la - prev_t) / la, 3)})
            last.pop("inferred_length", None)
        new_p.append(last)
        new_n.append(b_lab)
    # ② 두 끝에 세로관 · 사이를 올린다
    pos = {lab: j for j, lab in enumerate(new_n)}
    verts = []
    for idx in range(len(pairs)):
        A, B = tag_at.get(("a", idx)), tag_at.get(("b", idx))
        if A is None or B is None or pos[B] <= pos[A]:
            report["pairs"] -= 1
            report["unpaired"] += 2
            continue
        ja, jb = pos[A], pos[B]
        z0 = float(nodes[A].get("elevation") or 0.0)
        top_a = new_node(_xy(nodes, A), z0 + rise, A)
        top_b = new_node(_xy(nodes, B), z0 + rise, B)
        for lab in new_n[ja + 1:jb]:
            nodes[lab]["elevation"] = round(float(nodes[lab].get("elevation") or 0.0) + rise, 3)
        new_p[ja]["in"] = top_a
        new_p[jb - 1]["out"] = top_b
        for like, a, b, dz in ((new_p[ja], A, top_a, rise), (new_p[jb - 1], top_b, B, -rise)):
            v = dict(like)
            v.update({"label": new_label(), "in": a, "out": b,
                      "length": round(rise, 3), "elev": round(dz, 3), "hop": True})
            v.pop("inferred_length", None)
            v.pop("inferred_elev", None)
            verts.append(v)
    result["pipes"] = new_p + verts
    got = _chain(result)
    if got is None:
        raise ValueError("호 우회를 넣은 뒤 경로가 한 줄로 이어지지 않습니다.")
    return got[0], got[1], report


# ── ⑤ 부속 ─────────────────────────────────────────────────────────────────
def _fittings(result, order_n, order_p, phys):
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.design.fitting import build_fittings
    from src.pipenet_converter.graph.fitting_policy import CALCULATION_FITTINGS
    nodes = {str(n["label"]): n for n in result["nodes"]}
    net = {"pipe_data": {str(p["label"]): {"start": str(p["in"]), "end": str(p["out"])}
                         for p in order_p}}
    node_xy = {lab: (float(n["x"]) / 1000.0, float(n["y"]) / 1000.0)
               for lab, n in nodes.items()}
    node_z = {lab: float(n.get("elevation") or 0.0) for lab, n in nodes.items()}
    bores = {str(p["label"]): (int(float(p["dia"])) if p.get("dia") not in (None, "") else None,
                               p.get("dia_source")) for p in order_p}
    parents = {order_n[i + 1]: order_n[i] for i in range(len(order_n) - 1)}
    deg_path = defaultdict(int)
    for p in order_p:
        deg_path[str(p["in"])] += 1
        deg_path[str(p["out"])] += 1
    ph = {lab: max(int(phys.get(lab, 0)), deg_path[lab]) for lab in nodes}
    for lab in (order_n[0], order_n[-1]):
        ph[lab] = deg_path[lab]               # 끝은 부속을 안 단다(평면도·기계실 몫)
    got = build_fittings(net, node_xy, bores, parents=parents, node_z=node_z,
                         phys=ph, allowed_kinds=CALCULATION_FITTINGS)
    rows = []
    for p in order_p:
        rec = (got.get("per_pipe") or {}).get(str(p["label"])) or {}
        p["eq_len"] = round(float(rec.get("equivalent_length") or 0.0), 3)
        for kind in rec.get("fittings") or ():
            rows.append({"pipe": str(p["label"]), "in": str(p["in"]),
                         "out": str(p["out"]), "type": kind, "count": "1"})
    return rows, {"counts": dict(got.get("counts") or {}),
                  "unresolved_kind": int(got.get("unresolved_kind") or 0),
                  "unresolved_length": int(got.get("unresolved_length") or 0),
                  "unresolved_kind_items": list(got.get("unresolved_kind_items") or ())}


def drawn_degree(graph, stats, labels_xy) -> dict:
    """경로 노드마다 «도면에 그려진» 연결 수 — 다리(추정 이음)는 빼고 센다."""
    bridges = set()
    for key in ("tolerance_bridge_edges", "forced_bridge_edges"):
        for (a, b) in ((stats or {}).get(key) or ()):
            bridges.add(frozenset(((int(a[0]), int(a[1])), (int(b[0]), int(b[1])))))
    where = {}
    for g in (graph or {}):
        where[(int(round(g[0])), int(round(g[1])))] = g
    out = {}
    for lab, (x, y) in labels_xy.items():
        g = where.get((int(round(x)), int(round(y))))
        if g is None:
            continue
        ig = (int(round(g[0])), int(round(g[1])))
        n = 0
        for m in graph.get(g, ()):
            im = (int(round(m[0])), int(round(m[1])))
            if frozenset((ig, im)) not in bridges:
                n += 1
        out[lab] = n
    return out


def tidy_same_level(result, *, entities, graph, stats, layers) -> dict:
    """«같은 층» 추출 결과를 평면 그대로 다듬는다(②~⑤). 보고는 `same_level` 칸."""
    rep = {"merged_nodes": 0, "av_tail": None,
           "hops": {"pairs": 0, "unpaired": 0}, "fittings": {}, "unresolved_kind": 0}
    got = _chain(result)
    if got is None:
        rep["skipped"] = "경로가 한 줄로 이어지지 않아 다듬지 않았습니다"
        result["same_level"] = rep
        return result
    order_n, order_p = got
    order_n, order_p, rep["merged_nodes"] = _merge_straight(result, order_n, order_p)
    order_n, order_p, rep["av_tail"] = _trim_av_tail(
        result, order_n, order_p, float((stats or {}).get("scale_ratio") or 1.0))
    nodes = {str(n["label"]): n for n in result["nodes"]}
    phys = drawn_degree(graph, stats, {lab: _xy(nodes, lab) for lab in order_n})
    lay = set(layers or (stats or {}).get("layer_filter_used") or ())
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.dto import BRANCH_DEFAULT_M
    tol = max(float((stats or {}).get("snap_eps_mm") or 0.0), 1.0)
    order_n, order_p, rep["hops"] = _hops(result, order_n, order_p,
                                          _arcs_on(entities, lay), phys,
                                          float(BRANCH_DEFAULT_M), tol)
    rows, fit = _fittings(result, order_n, order_p, phys)
    result["fittings"] = rows
    rep["fittings"] = fit["counts"]
    rep["unresolved_kind"] = fit["unresolved_kind"]
    rep["unresolved_kind_items"] = fit["unresolved_kind_items"]
    result["pipes"] = order_p
    result["path_node_count"] = len(order_n)
    result["total_pipe_length_m"] = sum(float(p["length"]) for p in order_p)
    result["same_level"] = rep
    return result

# -*- coding: utf-8 -*-
"""시제품 3·4단계 이음 소유 모듈 — 본체 poc6 로직 복사.

실행 경로에서 `_tmp_dxf_extract_poc6` 를 부르지 않는다.
SNAP·격자·각도·묶음키는 stage1 공유 import.
"""
import math
from collections import defaultdict

from services.cad_import.pipeline.stage1 import SNAP, _grid_put, _grid_near, _eb_key, ang_between


def _default_knobs(*names) -> dict:
    """이 단계가 쓰는 눈금 기본값 — **stage1 의 DEFAULT_KNOBS 에서 가져온다.**

    ★여기 숫자를 다시 적지 않는다. 종전에는 `r1_cand=2000.0` 이 적혀 있었는데
      본체는 2026-08-03 에 1,500 으로 내렸다(구 2,000 이 1,649·1,861mm 오이음을
      냈다). 지금은 부르는 쪽이 늘 `knobs` 를 넘겨 덮으므로 잠복 상태였지만,
      `knobs=None` 으로 부르는 자리가 하나 생기는 순간 그 사고가 되살아난다.
    """
    from services.cad_import.pipeline.stage1 import DEFAULT_KNOBS
    want = names or ("r1_cand", "r1_meas_cap", "r1_lat_tol",
                     "r1_head_slack", "r1_cover_slack")
    return {k: DEFAULT_KNOBS[k] for k in want}

_INDEX_CELL = 500.0


def _bbox_put(grid, idx, a, b, pad=0.0, cell=_INDEX_CELL):
    x0, x1 = min(a[0], b[0]) - pad, max(a[0], b[0]) + pad
    y0, y1 = min(a[1], b[1]) - pad, max(a[1], b[1]) + pad
    for gx in range(int(x0 // cell), int(x1 // cell) + 1):
        for gy in range(int(y0 // cell), int(y1 // cell) + 1):
            grid[(gx, gy)].append(idx)


def _bbox_ids(grid, a, b, pad=0.0, cell=_INDEX_CELL):
    """bbox 후보를 원래 행 순서로 반환한다."""
    x0, x1 = min(a[0], b[0]) - pad, max(a[0], b[0]) + pad
    y0, y1 = min(a[1], b[1]) - pad, max(a[1], b[1]) + pad
    found = set()
    for gx in range(int(x0 // cell), int(x1 // cell) + 1):
        for gy in range(int(y0 // cell), int(y1 // cell) + 1):
            found.update(grid.get((gx, gy), ()))
    return sorted(found)


def _arr_edge(g, m, i, ux, uy):
    """노드 i에 방향(ux,uy)으로 공선 도착(≤8°)하는 최적 이웃(먼 끝) — 없으면
    None. 그 방향으로 계속되는 간선(≤30°)이 있으면 '도착 후 정지'가 아니므로
    None. 각도 눈금은 본체 이음⑦(arrives_and_stops)과 동일."""
    best = None
    xi, yi = g.pts[i]
    for k in m.get(i, ()):
        wx, wy = g.pts[k][0] - xi, g.pts[k][1] - yi
        if ang_between((wx, wy), (ux, uy)) <= 30.0:
            return None
        a = ang_between((-wx, -wy), (ux, uy))
        if a <= 8.0 and (best is None or a < best[0]):
            best = (a, k)
    return best[1] if best else None


def join_by_head_cover(g, ebundle, heads, knobs=None):
    """헤드 원이 직선 틈을 설명하면 이음 [사다리 D / 헤드걸침].

    조건 둘 다: ① 원이 틈 축을 자를 것(횡이탈 < 반지름)
                ② 원 테두리 밖으로 남는 길이가 양 끝 각각 r1_cover_slack 이하
    즉 '관이 헤드 원을 관통하다가 원 자리만 비워져 있다'는 모양만 잇는다.
    ★[2026-08-02] 눈금을 원 크기 비례(틈 ≤ 4r · 남는 길이 ≤ 2r)로 바꿨다가
      되돌렸다. 비례 눈금은 남는 길이를 360mm까지 봐주므로, 헤드가 틈 한쪽
      끝에 딱 붙어 있고 반대쪽 257mm가 빈 자리인 모양까지 통과시킨다. MF3
      (894428~895045, y=460780·464680·471730·475630) 네 곳이 그렇게 잘못
      이어졌다 — 거기 헤드는 관을 걸친 게 아니라 자기 접속관 하나만 동쪽에
      달린 '가지 끝 헤드'였고, 이음은 서로 다른 두 헤드의 접속관을 빈 자리
      617mm를 건너 붙여 좌·우 가지관을 사실상 연결해 버렸다.
      대신 눈금 숫자만 120 → 150mm 로 올렸다. 판정에 쓰는 남는 길이는
      원 반지름이 아니라 현(弦) 기준이므로 헤드가 축에서 비껴 있을수록 커진다
      — MF3 겹친 헤드 실측 125·134 / 92·119 / 121·101mm 로 120mm 를 근소하게
      넘어 셋 중 하나만 이어지던 것이 원인이었다. 반대로 위 잘못된 네 곳은
      257mm 라 150mm 로도 확실히 막힌다.
    heads = [(cx, cy, r), ...]  또는 [(cx, cy, r, kind), ...]
    반환: 이음 목록.
    """
    kn = _default_knobs()
    if knobs:
        kn.update(knobs)
    disks = []
    for h in heads:
        if len(h) >= 3 and h[2] > 0:
            disks.append((float(h[0]), float(h[1]), float(h[2])))
    m = g.adj()
    node_ids = [i for i in range(len(g.pts)) if m.get(i)]
    R = kn["r1_cand"]
    ngrid = defaultdict(list)
    for i in node_ids:
        x, y = g.pts[i]
        _grid_put(ngrid, R, x, y, i)

    def head_explains(pa, pb):
        dx, dy = pb[0] - pa[0], pb[1] - pa[1]
        L = math.hypot(dx, dy)
        if L < 1e-9:
            return None
        ux, uy = dx / L, dy / L
        best = None
        side = kn["r1_cover_slack"]
        for (qx, qy, qr) in disks:
            wx, wy = qx - pa[0], qy - pa[1]
            along = wx * ux + wy * uy
            lat = abs(-wx * uy + wy * ux)
            if lat >= qr:                 # ① 원이 축을 자름
                continue
            half = math.sqrt(qr * qr - lat * lat)
            # ② 원 밖으로 남는 길이 = 그리기 여유 수준이어야 한다
            if along - half <= side and L - (along + half) <= side:
                sc = (lat, abs(along - L / 2))
                if best is None or sc < best[0]:
                    best = (sc, qx, qy, qr, lat)
        return None if best is None else best[1:]

    elig = []
    for i in node_ids:
        xi, yi = g.pts[i]
        for j in _grid_near(ngrid, R, xi, yi):
            if j <= i:
                continue
            xj, yj = g.pts[j]
            d = math.hypot(xj - xi, yj - yi)
            if not (SNAP < d <= R):
                continue
            ux, uy = (xj - xi) / d, (yj - yi) / d
            ki = _arr_edge(g, m, i, ux, uy)
            if ki is None:
                continue
            kj = _arr_edge(g, m, j, -ux, -uy)
            if kj is None:
                continue
            A1, B2 = g.pts[ki], g.pts[kj]
            vx, vy = B2[0] - A1[0], B2[1] - A1[1]
            L = math.hypot(vx, vy)
            if L < 1e-9:
                continue
            rx, ry = vx / L, vy / L
            dev = max(abs(-(xi - A1[0]) * ry + (yi - A1[1]) * rx),
                      abs(-(xj - A1[0]) * ry + (yj - A1[1]) * rx))
            if dev > kn["r1_lat_tol"]:
                continue
            tA = (xi - A1[0]) * rx + (yi - A1[1]) * ry
            tB = (xj - A1[0]) * rx + (yj - A1[1]) * ry
            if not (1.0 <= tA and tA + 1.0 <= tB and tB <= L - 1.0):
                continue
            bi = ebundle.get(_eb_key(i, ki))
            bj = ebundle.get(_eb_key(j, kj))
            if bi is None or bi != bj:
                continue
            hit = head_explains((xi, yi), (xj, yj))
            if hit is None:
                continue
            hx, hy, hr, lat = hit
            elig.append(dict(i=i, j=j, d=d, dev=dev, bnd=bi,
                             hx=hx, hy=hy, hr=hr, lat=lat))

    best_of = {}
    for c in elig:
        sc = (c["dev"], c["d"])
        for end in (c["i"], c["j"]):
            if end not in best_of or sc < best_of[end][0]:
                best_of[end] = (sc, c)

    joins = []
    for c in elig:
        if best_of[c["i"]][1] is not c or best_of[c["j"]][1] is not c:
            continue
        if (min(c["i"], c["j"]), max(c["i"], c["j"])) in g.edges:
            continue
        g.add_edge(c["i"], c["j"])
        ebundle[_eb_key(c["i"], c["j"])] = c["bnd"]
        joins.append({
            "kind": "이음",
            "a": [round(g.pts[c["i"]][0], 1), round(g.pts[c["i"]][1], 1)],
            "b": [round(g.pts[c["j"]][0], 1), round(g.pts[c["j"]][1], 1)],
            "gap": round(c["d"], 1),
            "reason": ["헤드걸침잇기"],
            "sym_r": round(c["hr"], 1),
            "lat": round(c["lat"], 1),
            "head": [round(c["hx"], 1), round(c["hy"], 1)],
            "_nodes": (c["i"], c["j"]),
        })
    return joins


def join_by_through_main(g, ebundle, symbols, knobs=None, texts=None,
                         mat_layers=None, explain_segs=None):
    """끊어그린배관잇기 — 빈 직선 틈을 설명자가 있으면 이음.

    표시 용어(2026-08-02): 끊어그린배관잇기 (구 관행틈 / 통과메인).

    사다리 3단계(2026-08-01 로직 · 08-02 문자설명):
      · 기호 없는 마주 보는 직선 끊김
      · 틈 구간에 같은 축·같은 묶음 재료가 없을 것 (가짜 긴 tip-tip 차단)
      · 설명자: (a) 재료 통과선이 틈을 가로지름 또는
               (b) 재료 레이어 문자가 틈 축 위에 있음 또는
               (c) 비재료 설명자 배관(소화전 등)이 틈을 가로지름
                   — 2026-08-02 오너 합의 · 같은 날 오후 정식 반영(MF3)
    T·엘보·재부착은 하지 않는다.
    symbols = [(cx, cy, r, kind), ...] — 틈 자리 기호/헤드면 후보 제외.
    texts = [(x, y, h, ly), ...] 또는 World.texts 행 (ly,c,x,y,h,s)
    반환: (이음 목록, 부수정보 dict).
    """
    kn = _default_knobs("r1_cand", "r1_meas_cap", "r1_lat_tol",
                        "r1_sym_lat", "r1_sym_slack", "r1_head_slack")
    if knobs:
        kn.update(knobs)
    mat_layers = set(mat_layers or ())
    text_pts = []
    for t in (texts or ()):
        if len(t) >= 6:          # World.texts: ly,c,x,y,h,s
            ly, _c, x, y, h, _s = t[:6]
        elif len(t) >= 4:        # (x,y,h,ly)
            x, y, h, ly = t[:4]
        else:
            continue
        if h is None or h <= 0:
            continue
        if mat_layers and ly not in mat_layers:
            continue
        text_pts.append((float(x), float(y), float(h), ly))
    # 비재료 설명자 배관 (소화전 등) — 망에는 없고 틈 가로지름만 증거
    expl = []
    for seg in (explain_segs or ()):
        if len(seg) >= 2:
            a, b = seg[0], seg[1]
            if (math.hypot(b[0] - a[0], b[1] - a[1]) > 1.0):
                expl.append((a, b))
    m = g.adj()
    node_ids = [i for i in range(len(g.pts)) if m.get(i)]
    R = kn["r1_cand"]
    ngrid = defaultdict(list)
    for i in node_ids:
        x, y = g.pts[i]
        _grid_put(ngrid, R, x, y, i)

    # 통과선 탐색용 간선 목록 (한 번만)
    edge_segs = []
    for (ei, ej) in g.edges:
        a, b = g.pts[ei], g.pts[ej]
        edge_segs.append((ei, ej, a, b, ebundle.get(_eb_key(ei, ej))))
    edge_grid = defaultdict(list)
    edge_pad = max(1.0, float(kn["r1_lat_tol"]))
    for idx, (_ei, _ej, a, b, _eb) in enumerate(edge_segs):
        _bbox_put(edge_grid, idx, a, b, pad=edge_pad)

    symbol_grid = defaultdict(list)
    symbol_pad = max(
        [float(kn["r1_sym_lat"]), float(kn["r1_sym_slack"])]
        + [float(qr) + float(kn["r1_head_slack"])
           for _qx, _qy, qr, _k in symbols])
    for idx, (qx, qy, _qr, _k) in enumerate(symbols):
        _grid_put(symbol_grid, _INDEX_CELL, qx, qy, idx)

    text_grid = defaultdict(list)
    text_pad = max((max(float(h), 50.0) for _x, _y, h, _ly in text_pts),
                   default=0.0)
    for idx, (qx, qy, _h, _ly) in enumerate(text_pts):
        _grid_put(text_grid, _INDEX_CELL, qx, qy, idx)

    expl_grid = defaultdict(list)
    for idx, (a, b) in enumerate(expl):
        _bbox_put(expl_grid, idx, a, b)

    def point_ids(grid, pa, pb, pad):
        x0, x1 = min(pa[0], pb[0]) - pad, max(pa[0], pb[0]) + pad
        y0, y1 = min(pa[1], pb[1]) - pad, max(pa[1], pb[1]) + pad
        found = set()
        for gx in range(int(x0 // _INDEX_CELL), int(x1 // _INDEX_CELL) + 1):
            for gy in range(int(y0 // _INDEX_CELL), int(y1 // _INDEX_CELL) + 1):
                found.update(grid.get((gx, gy), ()))
        return sorted(found)

    def gap_has_symbol(pa, pb):
        dx, dy = pb[0] - pa[0], pb[1] - pa[1]
        L = math.hypot(dx, dy)
        if L < 1e-9:
            return True
        ux, uy = dx / L, dy / L
        for idx in point_ids(symbol_grid, pa, pb, symbol_pad):
            qx, qy, qr, _k = symbols[idx]
            wx, wy = qx - pa[0], qy - pa[1]
            along = wx * ux + wy * uy
            lat = abs(-wx * uy + wy * ux)
            if lat <= kn["r1_sym_lat"] and \
                    -kn["r1_sym_slack"] <= along <= L + kn["r1_sym_slack"]:
                return True
            if qr > 0 and lat < qr:
                half = math.sqrt(qr * qr - lat * lat)
                if along - half <= kn["r1_head_slack"] and \
                        L - (along + half) <= kn["r1_head_slack"]:
                    return True
        return False

    def gap_has_body(pa, pb, bnd):
        """틈 열린 구간에 같은 축·같은 묶음 재료가 있으면 True (1491형)."""
        dx, dy = pb[0] - pa[0], pb[1] - pa[1]
        L = math.hypot(dx, dy)
        if L < 1e-9:
            return True
        ux, uy = dx / L, dy / L
        margin = max(1.0, kn["r1_lat_tol"])
        for idx in _bbox_ids(edge_grid, pa, pb, pad=margin):
            _ei, _ej, a, b, eb = edge_segs[idx]
            if eb != bnd:
                continue
            lats = []
            alongs = []
            for p in (a, b):
                wx, wy = p[0] - pa[0], p[1] - pa[1]
                lats.append(abs(-wx * uy + wy * ux))
                alongs.append(wx * ux + wy * uy)
            if max(lats) > kn["r1_lat_tol"]:
                continue
            lo, hi = min(alongs), max(alongs)
            if hi > margin and lo < L - margin:
                return True
        return False

    def gap_has_text(pa, pb):
        """틈 축 위 재료 레이어 문자 — 표기 자리로 끊어 그린 증거."""
        if not text_pts:
            return None
        dx, dy = pb[0] - pa[0], pb[1] - pa[1]
        L = math.hypot(dx, dy)
        if L < 1e-9:
            return None
        ux, uy = dx / L, dy / L
        for idx in point_ids(text_grid, pa, pb, text_pad):
            qx, qy, h, ly = text_pts[idx]
            wx, wy = qx - pa[0], qy - pa[1]
            along = wx * ux + wy * uy
            lat = abs(-wx * uy + wy * ux)
            slack = max(h, 50.0)
            # 횡이탈: 글자높이만큼(관 옆 표기). 0.6·h는 BF4 45341 문자가
            # 45158 축에서 183mm로 3mm 차이 탈락 → h로 둔다.
            lat_tol = max(kn["r1_sym_lat"], h)
            if -slack <= along <= L + slack and lat <= lat_tol:
                return (qx, qy, h, ly)
        return None

    def through_crosser(pa, pb, tip_i, tip_j):
        """틈을 가로지르는 통과선 1개라도 있으면 (교차점, 끝점쌍) 아니면 None.

        통과 = 메인과 평행하지 않고, 교차점이 틈 안·통과선 몸통(끝 아님).
        가지 끝이 메인에 닿는 T는 통과선 끝점이라 제외된다.
        """
        dx, dy = pb[0] - pa[0], pb[1] - pa[1]
        L = math.hypot(dx, dy)
        if L < 1e-9:
            return None
        ux, uy = dx / L, dy / L
        tip_eps = 30.0
        cos_par = math.cos(math.radians(25.0))
        best = None
        for idx in _bbox_ids(edge_grid, pa, pb):
            ei, ej, a, b, _eb = edge_segs[idx]
            if ei in (tip_i, tip_j) or ej in (tip_i, tip_j):
                continue
            vx, vy = b[0] - a[0], b[1] - a[1]
            Lv = math.hypot(vx, vy)
            if Lv < 1e-9:
                continue
            if abs(ux * vx + uy * vy) / Lv >= cos_par:
                continue
            # pa + t*(pb-pa) = a + s*(b-a)
            den = dx * (-vy) - dy * (-vx)
            if abs(den) < 1e-12:
                continue
            t = ((a[0] - pa[0]) * (-vy) - (a[1] - pa[1]) * (-vx)) / den
            s = ((a[0] - pa[0]) * (-dy) - (a[1] - pa[1]) * (-dx)) / den
            if not (0.02 <= t <= 0.98):
                continue
            along_mm = s * Lv
            if not (tip_eps <= along_mm <= Lv - tip_eps):
                continue
            hx = pa[0] + dx * t
            hy = pa[1] + dy * t
            sc = abs(t - 0.5)
            if best is None or sc < best[0]:
                best = (sc, hx, hy, a, b)
        if best is None:
            return None
        return (best[1], best[2], best[3], best[4])


    def explain_crosser(pa, pb):
        """비재료 설명자 배관이 틈 내부를 가로지르면 교차점."""
        if not expl:
            return None
        dx, dy = pb[0] - pa[0], pb[1] - pa[1]
        L = math.hypot(dx, dy)
        if L < 1e-9:
            return None
        ux, uy = dx / L, dy / L
        tip_eps = 30.0
        cos_par = math.cos(math.radians(25.0))
        best = None
        for idx in _bbox_ids(expl_grid, pa, pb):
            a, b = expl[idx]
            vx, vy = b[0] - a[0], b[1] - a[1]
            Lv = math.hypot(vx, vy)
            if Lv < 1e-9:
                continue
            if abs(ux * vx + uy * vy) / Lv >= cos_par:
                continue
            den = dx * (-vy) - dy * (-vx)
            if abs(den) < 1e-12:
                continue
            t = ((a[0] - pa[0]) * (-vy) - (a[1] - pa[1]) * (-vx)) / den
            s = ((a[0] - pa[0]) * (-dy) - (a[1] - pa[1]) * (-dx)) / den
            if not (0.02 <= t <= 0.98):
                continue
            along_mm = s * Lv
            if not (tip_eps <= along_mm <= Lv - tip_eps):
                continue
            hx = pa[0] + dx * t
            hy = pa[1] + dy * t
            sc = abs(t - 0.5)
            if best is None or sc < best[0]:
                best = (sc, hx, hy, a, b)
        if best is None:
            return None
        return (best[1], best[2], best[3], best[4])

    cands = []
    n_sym = n_body = n_nocross = n_text = n_explain = 0
    for i in node_ids:
        xi, yi = g.pts[i]
        for j in _grid_near(ngrid, R, xi, yi):
            if j <= i:
                continue
            xj, yj = g.pts[j]
            d = math.hypot(xj - xi, yj - yi)
            if not (SNAP < d <= R):
                continue
            ux, uy = (xj - xi) / d, (yj - yi) / d
            ki = _arr_edge(g, m, i, ux, uy)
            if ki is None:
                continue
            kj = _arr_edge(g, m, j, -ux, -uy)
            if kj is None:
                continue
            A1, B2 = g.pts[ki], g.pts[kj]
            vx, vy = B2[0] - A1[0], B2[1] - A1[1]
            Larm = math.hypot(vx, vy)
            if Larm < 1e-9:
                continue
            rx, ry = vx / Larm, vy / Larm
            dev = max(abs(-(xi - A1[0]) * ry + (yi - A1[1]) * rx),
                      abs(-(xj - A1[0]) * ry + (yj - A1[1]) * rx))
            if dev > kn["r1_meas_cap"]:
                continue
            tA = (xi - A1[0]) * rx + (yi - A1[1]) * ry
            tB = (xj - A1[0]) * rx + (yj - A1[1]) * ry
            if not (1.0 <= tA and tA + 1.0 <= tB and tB <= Larm - 1.0):
                continue
            if dev > kn["r1_lat_tol"]:
                continue
            bi = ebundle.get(_eb_key(i, ki))
            bj = ebundle.get(_eb_key(j, kj))
            if bi is None or bi != bj:
                continue
            pa, pb = (xi, yi), (xj, yj)
            if gap_has_symbol(pa, pb):
                n_sym += 1
                continue
            if gap_has_body(pa, pb, bi):
                n_body += 1
                continue
            hit = through_crosser(pa, pb, i, j)
            how = "통과선"
            cross_at = None
            cross_ab = None
            text_hit = None
            if hit is not None:
                cross_at = (hit[0], hit[1])
                cross_ab = (hit[2], hit[3])
            else:
                ehit = explain_crosser(pa, pb)
                if ehit is not None:
                    how = "설명자"
                    n_explain += 1
                    cross_at = (ehit[0], ehit[1])
                    cross_ab = (ehit[2], ehit[3])
                else:
                    text_hit = gap_has_text(pa, pb)
                    if text_hit is None:
                        n_nocross += 1
                        continue
                    how = "문자"
                    n_text += 1
                    cross_at = (text_hit[0], text_hit[1])
            cands.append(dict(i=i, j=j, d=d, dev=dev, bnd=bi, how=how,
                              cross_at=cross_at, cross_ab=cross_ab,
                              text=text_hit))

    best_of = {}
    for c in cands:
        sc = (c["dev"], c["d"])
        for end in (c["i"], c["j"]):
            if end not in best_of or sc < best_of[end][0]:
                best_of[end] = (sc, c)

    # ★같은 덩어리 이음 금지 [2026-08-05 D11 종결] — 이음 양 끝이 이미 한
    #   덩어리면 잇지 않는다. 이 규칙은 도면에 없는 관(고리)을 만든다:
    #   apt 실측 5/5 — 후렉시블 두 가닥이 이미 같은 분홍 관에 붙어 있는데
    #   그 중간을 45도 대각선 1,042mm 로 또 이었다(덩어리 수 불변 = 순수 고리).
    #   ★D9(색 장벽 폐지)를 먼저 고쳐 한 덩어리가 되고 나서야 이 조건이
    #   가능해졌다 — 그 전에는 셋으로 갈라져 보여 '같은 덩어리'가 아니었다.
    #   막은 자리는 조용히 버리지 않고 skipped(경합 목록)에 남긴다.
    _par = {}

    def _find(x):
        _par.setdefault(x, x)
        while _par[x] != x:
            _par[x] = _par[_par[x]]
            x = _par[x]
        return x

    for (i2, j2) in g.edges:
        ri, rj = _find(i2), _find(j2)
        if ri != rj:
            _par[ri] = rj
    joins, skipped = [], []
    for c in cands:
        if best_of[c["i"]][1] is not c or best_of[c["j"]][1] is not c:
            skipped.append(c)
            continue
        if (min(c["i"], c["j"]), max(c["i"], c["j"])) in g.edges:
            continue
        if _find(c["i"]) == _find(c["j"]):
            skipped.append(dict(c, why="같은덩어리(고리 금지)"))
            continue
        g.add_edge(c["i"], c["j"])
        _par[_find(c["i"])] = _find(c["j"])
        ebundle[_eb_key(c["i"], c["j"])] = c["bnd"]
        rec = {
            "kind": "이음",
            "a": [round(g.pts[c["i"]][0], 1), round(g.pts[c["i"]][1], 1)],
            "b": [round(g.pts[c["j"]][0], 1), round(g.pts[c["j"]][1], 1)],
            "gap": round(c["d"], 1),
            "reason": ["끊어그린배관잇기", c["how"]],
            "dev": round(c["dev"], 2),
            "_nodes": (c["i"], c["j"]),
        }
        if c["cross_at"] is not None:
            rec["cross"] = [round(c["cross_at"][0], 1),
                            round(c["cross_at"][1], 1)]
        joins.append(rec)
    side = {
        "n_cand": len(cands),
        "n_join": len(joins),
        "n_skipped": len(skipped),
        "n_sym_skip": n_sym,
        "n_body_skip": n_body,
        "n_nocross": n_nocross,
        "n_text": n_text,
        "n_explain": n_explain,
        "skipped": skipped,
        "cands": cands,
    }
    return joins, side

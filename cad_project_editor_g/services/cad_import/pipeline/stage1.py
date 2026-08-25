# -*- coding: utf-8 -*-
"""시제품 1단계 소유 모듈 — 본체 poc3/poc6 로직 복사.

실행 경로에서 `_tmp_dxf_extract_poc6` / `_tmp_dxf_extract_poc3` 를 부르지 않는다.
이음 규칙은 바꾸지 않았다. 동작 동일 복제만.
"""
import math
import os
from collections import Counter, defaultdict

# ---- poc3 에서 복사 (SNAP · Graph · 기하) ----
SNAP = 30.0

# ---------------------------------------------------------------- 기하 공용
def seg_geom(a, b, x, y):
    (sx, sy), (tx, ty) = a, b
    dx, dy = tx - sx, ty - sy
    L2 = dx * dx + dy * dy
    if L2 <= 0:
        return math.hypot(x - sx, y - sy), 0.0, (sx, sy)
    u = ((x - sx) * dx + (y - sy) * dy) / L2
    uc = max(0.0, min(1.0, u))
    fx, fy = sx + uc * dx, sy + uc * dy
    return math.hypot(fx - x, fy - y), uc, (fx, fy)

def ang_between(v1, v2):
    n1 = math.hypot(*v1)
    n2 = math.hypot(*v2)
    if n1 < 1e-9 or n2 < 1e-9:
        return 0.0
    c = max(-1.0, min(1.0, (v1[0] * v2[0] + v1[1] * v2[1]) / (n1 * n2)))
    return math.degrees(math.acos(c))

# ---------------------------------------------------------------- 관망 그래프
class Graph:
    def __init__(self):
        self.pts = []          # 노드 좌표
        self.edges = set()     # (i,j) i<j
        self._grid = {}

    def _key(self, x, y):
        return (int(x // SNAP), int(y // SNAP))

    def node(self, x, y):
        kx, ky = self._key(x, y)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for idx in self._grid.get((kx + dx, ky + dy), ()):
                    px, py = self.pts[idx]
                    if math.hypot(px - x, py - y) <= SNAP:
                        return idx
        idx = len(self.pts)
        self.pts.append((x, y))
        self._grid.setdefault((kx, ky), []).append(idx)
        return idx

    def force_node(self, x, y):
        """스냅 없이 좌표 그대로 노드를 만든다.

        짧은 실배관 양 끝이 SNAP(30mm) 안에 있어 node()가 한 점으로
        접을 때, 접히지 않은 쪽 끝을 살릴 때만 쓴다.
        """
        x, y = float(x), float(y)
        idx = len(self.pts)
        self.pts.append((x, y))
        self._grid.setdefault(self._key(x, y), []).append(idx)
        return idx

    def add_edge(self, i, j):
        if i != j:
            self.edges.add((min(i, j), max(i, j)))

    def adj(self):
        m = defaultdict(set)
        for i, j in self.edges:
            m[i].add(j)
            m[j].add(i)
        return m

    def free_ends(self):
        m = self.adj()
        return [i for i in range(len(self.pts)) if len(m.get(i, ())) == 1], m

    def out_dir(self, i, m):
        j = next(iter(m[i]))
        px, py = self.pts[j]
        x, y = self.pts[i]
        return (x - px, y - py)

    def split_edge(self, i, j, fx, fy):
        # _grid 등록(2026-07-26): 예전에는 pts에만 append해서 이 노드가 좌표
        # 색인에 없었다(불변식 위반 — 뒤 규칙의 g.node()가 못 찾는다). 등록으로
        # 바로잡되, **측정 결과 6도면 추출 로그가 완전히 동일했다(효과 0).**
        # ★MF2 좌표 중복 노드의 원인은 이것이 아니다 — '접속 해제' 규칙이
        #   같은 좌표에 노드를 새로 만드는 것이 의도된 장치다(2269행 부근).
        k = len(self.pts)
        self.pts.append((fx, fy))
        self._grid.setdefault(self._key(fx, fy), []).append(k)
        self.edges.discard((min(i, j), max(i, j)))
        self.add_edge(i, k)
        self.add_edge(k, j)
        return k

# ---- poc6 에서 복사 (1단계 재료·그래프) ----
_APP_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
DWG_DIR = os.path.join(_APP_ROOT, "docs", "import", "DWG")

# ================================================================ 1. 읽기
# 62MB급 파일이 ezdxf 통째 로딩에서 메모리로 죽는 실측(2026-07-30) →
# 스트리밍 파서. 본체 collect_segs와 같은 대상(LINE/LWPOLYLINE/POLYLINE)
# + 블록 해제 재료(BLOCKS·INSERT·CIRCLE·ARC). 검산: MF 5레이어 6,042선 일치.
def _pairs(path):
    with open(path, encoding="utf-8", errors="replace") as f:
        while True:
            c = f.readline()
            if not c:
                return
            v = f.readline()
            if not v:
                return
            try:
                code = int(c.strip())
            except ValueError:
                continue
            yield code, v.rstrip("\r\n")

def read_dxf(path):
    """반환: ltab(레이어표), ents(모델공간 원시), bdefs(블록 정의)."""
    ltab = {}          # name -> [color, hidden]
    ents = []          # 원시 엔티티(dict) — ENTITIES 구간
    bdefs = defaultdict(list)   # name -> [ent dict]
    section = None
    state = 0          # 1=TABLES 2=ENTITIES 3=BLOCKS
    cur = None
    poly = None
    cur_block = None

    def close(ent, into):
        if not ent:
            return
        t = ent["t"]
        if t == "VERTEX":
            return
        if t in ("LINE", "LWPOLYLINE", "POLYLINE", "CIRCLE", "ARC", "INSERT",
                 "TEXT", "MTEXT"):
            into.append(ent)

    for code, val in _pairs(path):
        if code == 0:
            # 진행 중 엔티티 마감
            if state == 2:
                if val == "VERTEX" and poly is not None:
                    # 이전 VERTEX를 먼저 넣고 새 VERTEX로 넘어간다.
                    if (cur is not None and cur.get("t") == "VERTEX"
                            and "10" in cur and "20" in cur):
                        poly.setdefault("pts", []).append((cur["10"], cur["20"]))
                    cur = {"t": "VERTEX"}
                    continue
                if val == "SEQEND" and poly is not None:
                    # SEQEND 직전 VERTEX — 빼먹으면 2점 원(SP01-04)이 증발.
                    if (cur is not None and cur.get("t") == "VERTEX"
                            and "10" in cur and "20" in cur):
                        poly.setdefault("pts", []).append((cur["10"], cur["20"]))
                    close(poly, ents)
                    poly = None
                    cur = None
                    continue
                if cur is not None and cur is not poly:
                    close(cur, ents)
            elif state == 3 and cur_block is not None:
                if val == "VERTEX" and poly is not None:
                    if (cur is not None and cur.get("t") == "VERTEX"
                            and "10" in cur and "20" in cur):
                        poly.setdefault("pts", []).append((cur["10"], cur["20"]))
                    cur = {"t": "VERTEX"}
                    continue
                if val == "SEQEND" and poly is not None:
                    if (cur is not None and cur.get("t") == "VERTEX"
                            and "10" in cur and "20" in cur):
                        poly.setdefault("pts", []).append((cur["10"], cur["20"]))
                    close(poly, bdefs[cur_block])
                    poly = None
                    cur = None
                    continue
                if cur is not None and cur is not poly and cur["t"] != "@BLKHDR":
                    close(cur, bdefs[cur_block])
            if val == "SECTION":
                section = "?"
                cur = None
            elif val == "ENDSEC":
                state = 0
                section = None
                cur = None
                poly = None
            elif state == 3:
                if val == "BLOCK":
                    cur = {"t": "@BLKHDR"}
                elif val == "ENDBLK":
                    cur_block = None
                    cur = None
                elif val == "POLYLINE":
                    poly = {"t": "POLYLINE"}
                    cur = poly
                else:
                    cur = {"t": val}
            elif state == 2:
                if val == "POLYLINE":
                    poly = {"t": "POLYLINE"}
                    cur = poly
                else:
                    cur = {"t": val}
            elif state == 1 and val == "LAYER":
                cur = {"t": "@LAYER"}
            else:
                cur = None
            continue
        if code == 2 and section == "?":
            section = val
            state = 1 if val == "TABLES" else 2 if val == "ENTITIES" else \
                3 if val == "BLOCKS" else 0
            continue
        if cur is None:
            continue
        t = cur["t"]
        if t == "@LAYER":
            if code == 2:
                cur["name"] = val
                ltab.setdefault(val, [7, False])
            elif code == 62 and "name" in cur:
                try:
                    c62 = int(val)
                except ValueError:
                    c62 = 7
                ltab[cur["name"]][0] = abs(c62)
                if c62 < 0:                       # 음수 색 = 레이어 꺼짐
                    ltab[cur["name"]][1] = True
            elif code == 70 and "name" in cur:
                try:
                    if int(val) & 1:              # bit1 = 동결
                        ltab[cur["name"]][1] = True
                except ValueError:
                    pass
            continue
        if t == "@BLKHDR":
            if code == 2:
                cur_block = val
            continue
        if code == 2:
            cur["2"] = val
        elif code == 1:
            cur["1"] = val
        elif code == 3 and t == "MTEXT":
            cur["1"] = (cur.get("1") or "") + val
        elif code == 8:
            cur["8"] = val
        elif code in (62, 66, 67, 70):
            try:
                cur[str(code)] = int(val)
            except ValueError:
                pass
        elif code in (10, 20, 11, 21, 40, 41, 42, 50, 51):
            try:
                f = float(val)
            except ValueError:
                continue
            if t == "LWPOLYLINE" and code in (10, 20, 42):
                if code == 10:
                    cur["_x"] = f
                elif code == 20:
                    if "_x" in cur:
                        cur.setdefault("pts", []).append((cur.pop("_x"), f))
                        cur.setdefault("bul", []).append(0.0)
                elif cur.get("bul"):
                    # 42 = 방금 읽은 꼭짓점의 bulge — 그 점부터 다음 점까지가 호.
                    cur["bul"][-1] = f
            else:
                cur[str(code)] = f
    return ltab, ents, bdefs

# ================================================================ 2. 해제
# 정리서 §1: 블록 전부 해제(기하만, 깊이 2) — 한 세계. 블록 이름은 저장하지
# 않는다(명찰 폐기). 색 상속: 0(블록따름)=삽입물 색 · 256(레이어따름)=엔티티
# 레이어 색 · 블록 내부 레이어 '0'=삽입물 레이어. [실측 검증 1회 — §1-1]
class World:
    def __init__(self):
        self.segs = []      # (layer, color, a, b) — 해제 포함 전체
        self.raw_segs = []  # (layer, color, a, b) — 원시 선만(블록 획 제외)
        self.circles = []   # (layer, color, cx, cy, r)
        self.arcs = []      # (layer, color, cx, cy, r)
        self.arc_ang = []   # arcs 와 같은 순 · (sa, sweep) 또는 None
        self.texts = []     # (layer, color, x, y, h, s) — TEXT/MTEXT

def _bulge_arc(p, q, bg):
    """bulge 마디 → 원호 (cx, cy, r). 직선/영길이면 None. 좌표는 변환 전.

    AutoCAD 정의 bulge = tan(사잇각/4) — 양수 반시계 · 음수 시계.
      반지름 = 현·(1+b²)/(4|b|)
      중심   = 현 중점 + 왼쪽 법선 · 현·(1-b²)/(4b)      (b<0 이면 반대쪽)
    """
    (x1, y1), (x2, y2) = p, q
    c = math.hypot(x2 - x1, y2 - y1)
    if c < 1e-9 or abs(bg) < 1e-12:
        return None
    ux, uy = (x2 - x1) / c, (y2 - y1) / c
    h = c * (1.0 - bg * bg) / (4.0 * bg)
    return ((x1 + x2) / 2.0 - uy * h, (y1 + y2) / 2.0 + ux * h,
            c * (1.0 + bg * bg) / (4.0 * abs(bg)))

def _norm360(a):
    a = float(a) % 360.0
    return a + 360.0 if a < 0.0 else a

def _sweep_ccw(sa, ea):
    return (_norm360(ea) - _norm360(sa)) % 360.0

def _world_arc_ang(xf, lcx, lcy, lr, sa, ea):
    """로컬 시작/끝각 → 세계 sa, sweep. INSERT 회전 반영."""
    def lp(d):
        a = math.radians(d)
        return lcx + lr * math.cos(a), lcy + lr * math.sin(a)
    wcx, wcy = xf(lcx, lcy)
    sx, sy = xf(*lp(sa))
    ex, ey = xf(*lp(ea))
    sa_w = math.degrees(math.atan2(sy - wcy, sx - wcx))
    ea_w = math.degrees(math.atan2(ey - wcy, ex - wcx))
    sw0 = _sweep_ccw(sa, ea)
    mx, my = xf(*lp(sa + sw0 * 0.5))
    mid = math.degrees(math.atan2(my - wcy, mx - wcx))
    sw = _sweep_ccw(sa_w, ea_w)
    d = (_norm360(mid) - _norm360(sa_w)) % 360.0
    if sw > 0.0 and sw < 360.0 - 1e-12 and d > sw + 1e-12:
        sa_w, ea_w = ea_w, sa_w
        sw = _sweep_ccw(sa_w, ea_w)
    return _norm360(sa_w), sw

def explode(ltab, ents, bdefs):
    w = World()
    hidden = {n for n, (_c, h) in ltab.items() if h}
    hid_cnt = [0]

    def lcolor(lay):
        return ltab.get(lay, (7, False))[0]

    def resolve(lay, col, ctx_lay, ctx_col):
        lay2 = ctx_lay if (lay == "0" and ctx_lay) else lay
        if col is None or col == 256:
            c2 = lcolor(lay2)
        elif col == 0:
            c2 = ctx_col if ctx_col is not None else 7
        else:
            c2 = col
        return lay2, c2

    def put_arc(lay, col, cx, cy, r, sa=None, sweep=None):
        w.arcs.append((lay, col, cx, cy, r))
        if sa is None or sweep is None:
            w.arc_ang.append(None)
        else:
            w.arc_ang.append((float(sa), float(sweep)))

    def emit(ent, xf, ctx_lay, ctx_col, depth):
        t = ent["t"]
        lay, col = resolve(ent.get("8", "0"), ent.get("62"), ctx_lay, ctx_col)
        if lay in hidden:
            hid_cnt[0] += 1
            return
        if ent.get("67") == 1:
            return                                  # 페이퍼스페이스 제외
        if t == "LINE":
            if all(k in ent for k in ("10", "20", "11", "21")):
                a = xf(ent["10"], ent["20"])
                b = xf(ent["11"], ent["21"])
                if math.hypot(b[0] - a[0], b[1] - a[1]) > 1.0:
                    w.segs.append((lay, col, a, b))
                    if depth == 0:
                        w.raw_segs.append((lay, col, a, b))
        elif t in ("LWPOLYLINE", "POLYLINE"):
            loc = list(ent.get("pts", []))
            pts = [xf(*p) for p in loc]
            bul = ent.get("bul") or ()
            closed = bool(ent.get("70", 0) & 1)
            # 닫힌 2점 폴리라인 = 지름 양끝을 찍은 원(채운 원/도넛 관행).
            # 선분으로 풀면 짧은 직경 획만 남고 헤드로 안 잡힌다.
            if closed and len(pts) == 2:
                (x1, y1), (x2, y2) = pts
                cx, cy = (x1 + x2) / 2.0, (y1 + y2) / 2.0
                r = 0.5 * math.hypot(x2 - x1, y2 - y1)
                if r > 1.0:
                    w.circles.append((lay, col, cx, cy, r))
            else:
                order = list(range(len(pts) - 1))
                if closed and len(pts) >= 3:
                    order.append(len(pts) - 1)      # 닫힘 = 마지막→처음
                for i in order:
                    a, b = pts[i], pts[(i + 1) % len(pts)]
                    ln = math.hypot(b[0] - a[0], b[1] - a[1])
                    # bulge(그룹 42)가 붙은 마디는 «휜 마디»다 — ARC 와 같은
                    # 자격으로 원호에 낸다. 현(직선)은 배관으로 남기지 않는다.
                    bg = bul[i] if i < len(bul) else 0.0
                    if abs(bg) * ln / 2.0 > 1.0:    # 활 높이 1mm 넘으면 호
                        loc_a = loc[i]
                        loc_b = loc[(i + 1) % len(loc)]
                        arc = _bulge_arc(loc_a, loc_b, bg)
                        if arc is not None:
                            sa = math.degrees(math.atan2(
                                loc_a[1] - arc[1], loc_a[0] - arc[0]))
                            ea = math.degrees(math.atan2(
                                loc_b[1] - arc[1], loc_b[0] - arc[0]))
                            if bg < 0.0:
                                sa, ea = ea, sa
                            sa_w, sw = _world_arc_ang(
                                xf, arc[0], arc[1], arc[2], sa, ea)
                            acx, acy = xf(arc[0], arc[1])
                            put_arc(lay, col, acx, acy, arc[2] * xf.scale,
                                    sa_w, sw)
                            continue
                    if ln > 1.0:
                        w.segs.append((lay, col, a, b))
                        if depth == 0:
                            w.raw_segs.append((lay, col, a, b))
        elif t in ("CIRCLE", "ARC"):
            if all(k in ent for k in ("10", "20", "40")):
                cx, cy = xf(ent["10"], ent["20"])
                r = ent["40"] * xf.scale
                if t == "CIRCLE":
                    w.circles.append((lay, col, cx, cy, r))
                else:
                    sa, ea = ent.get("50"), ent.get("51")
                    if sa is None or ea is None:
                        put_arc(lay, col, cx, cy, r)
                    else:
                        sa_w, sw = _world_arc_ang(
                            xf, ent["10"], ent["20"], ent["40"], sa, ea)
                        put_arc(lay, col, cx, cy, r, sa_w, sw)
        elif t in ("TEXT", "MTEXT"):
            if "10" in ent and "20" in ent:
                cx, cy = xf(ent["10"], ent["20"])
                h = float(ent.get("40", 0.0) or 0.0) * xf.scale
                s = ent.get("1") or ""
                if h > 0:
                    w.texts.append((lay, col, cx, cy, h, s))
        elif t == "INSERT" and depth < 2:
            nm = ent.get("2")
            if nm not in bdefs:
                return
            ix, iy = ent.get("10", 0.0), ent.get("20", 0.0)
            rot = math.radians(ent.get("50", 0.0))
            sx = ent.get("41", 1.0)
            sy = ent.get("42", 1.0)
            cr, sr = math.cos(rot), math.sin(rot)
            base = xf

            class XF:
                scale = abs(sx) * base.scale

                @staticmethod
                def __call__(px, py):
                    x, y = px * sx, py * sy
                    return base(ix + x * cr - y * sr, iy + x * sr + y * cr)
            xf2 = XF()
            for be in bdefs[nm]:
                emit(be, xf2, lay, col, depth + 1)

    class ID:
        scale = 1.0

        @staticmethod
        def __call__(x, y):
            return (x, y)

    for e in ents:
        emit(e, ID(), None, None, 0)
    return w, hid_cnt[0]

def _grid_put(g, cell, x, y, v):
    g[(int(x // cell), int(y // cell))].append(v)

def _grid_near(g, cell, x, y, rings=1):
    cx, cy = int(x // cell), int(y // cell)
    for dx in range(-rings, rings + 1):
        for dy in range(-rings, rings + 1):
            yield from g.get((cx + dx, cy + dy), ())

DEFAULT_KNOBS = dict(
    rim_tol=50.0, arc_max=300.0, gate_r=60.0,
    cluster_gap=300.0, small_len=600.0, small_r=300.0,
    head_size_eps=5.0, head_size_rel=0.10,
    a1_lat=1.0, a1_scan=50.0,
    a2_rim_in=90.0,          # 기호 테두리 '안쪽' 허용 [오너 확정 2026-08-02]
    a2_rim_out=80.0,         #   40→90: MF3 치우친 반원 1곳(틈 360=지름) 회수.
    a2_gap_eps=5.0, a2_gap_min=3,
    # ★a3_cross_side(통과선 치우침 눈금)는 폐기 [2026-08-03 오너 확정].
    #   숫자로 막으려다 MF3에서 '이어야 할 6곳'을 같이 죽였다. 진짜 원인은
    #   기호 획(작대기·관말 캡)이 배관으로 잡힌 것 — symbol_strokes 로 해결.

    # ★틈 길이 상한 1,500mm [2026-08-03 오너 확정] — 구 2,000.
    #   확인된 정당한 이음의 최대가 1,241mm(MF3 상세도)·평면 736mm 이므로
    #   잘려나가는 정당 이음이 없고, 사고였던 1,649·1,861mm 는 기호 규칙에
    #   이어 두 겹으로 막힌다. (3F가 뚫렸던 원인은 이 상한이 아니라 폐기한
    #    '통과선 치우침 300mm' 규칙이었다.)
    r1_cand=1500.0, r1_tee=450.0, r1_sym_lat=60.0,
    r1_sym_slack=100.0, r1_meas_cap=60.0, r1_lat_tol=10.0,
    r1_head_slack=120.0, r1_cover_slack=150.0,
    r1_tee_sym_r=350.0, r1_tee_sym_lat=150.0,
    b1_touch=15.0, b2_cand=2000.0, b2_sym_r=300.0,
    b2_ang_lo=30.0, b2_ang_hi=150.0,
)

_HG_AXIS_ANG = 5.0       # 토막↔재료끝 같은 축 판정(도)

_HG_ATTACH_CAP = 500.0   # 접착 틈 상한(mm) — 실제 게이트는 '틈 안에 헤드'

_HG_MIN_CAND_L = 200.0   # 이보다 짧은 다른색 토막은 심볼·범례로 본다

_LOOKUP_CELL = 1000.0
_B1_CELL = 100.0


def _bbox_grid(segments, cell=_LOOKUP_CELL, pad=0.0):
    """선분 bbox 격자. 후보만 줄이고 최종 판정은 기존 거리식이 한다."""
    grid = defaultdict(list)
    pad = float(pad)
    for i, (a, b) in enumerate(segments):
        x0, x1 = min(a[0], b[0]) - pad, max(a[0], b[0]) + pad
        y0, y1 = min(a[1], b[1]) - pad, max(a[1], b[1]) + pad
        for gx in range(int(x0 // cell), int(x1 // cell) + 1):
            for gy in range(int(y0 // cell), int(y1 // cell) + 1):
                grid[(gx, gy)].append(i)
    return grid


def _bbox_near(grid, a, b, pad=0.0, cell=_LOOKUP_CELL):
    """선분 bbox(+pad)와 겹치는 격자 후보 인덱스."""
    x0, x1 = min(a[0], b[0]) - pad, max(a[0], b[0]) + pad
    y0, y1 = min(a[1], b[1]) - pad, max(a[1], b[1]) + pad
    seen = set()
    for gx in range(int(x0 // cell), int(x1 // cell) + 1):
        for gy in range(int(y0 // cell), int(y1 // cell) + 1):
            for i in grid.get((gx, gy), ()):
                if i not in seen:
                    seen.add(i)
                    yield i


def _point_grid(points, cell=_LOOKUP_CELL):
    grid = defaultdict(list)
    for i, row in enumerate(points):
        p = row[0]
        _grid_put(grid, cell, p[0], p[1], (i, row))
    return grid


def _circle_grid(circles, cell=_LOOKUP_CELL):
    grid = defaultdict(list)
    for row in circles:
        _grid_put(grid, cell, row[0], row[1], row)
    return grid


def _circles_near_segment(grid, a, b, pad, cell=_LOOKUP_CELL):
    """중심이 선분 bbox(+pad)에 있는 원 후보."""
    x0, x1 = min(a[0], b[0]) - pad, max(a[0], b[0]) + pad
    y0, y1 = min(a[1], b[1]) - pad, max(a[1], b[1]) + pad
    for gx in range(int(x0 // cell), int(x1 // cell) + 1):
        for gy in range(int(y0 // cell), int(y1 // cell) + 1):
            yield from grid.get((gx, gy), ())

def _hg_gap_has_head(p, q, heads):
    dx, dy = q[0] - p[0], q[1] - p[1]
    L = math.hypot(dx, dy)
    if L < 1e-9 or not heads:
        return False
    ux, uy = dx / L, dy / L
    for hx, hy, hr in heads:
        wx, wy = hx - p[0], hy - p[1]
        along = wx * ux + wy * uy
        lat = abs(-wx * uy + wy * ux)
        if lat <= max(hr, 30.0) and -hr <= along <= L + hr:
            return True
    return False

def _hg_axis_align(a, b, p, q, ang_tol=_HG_AXIS_ANG):
    sdx, sdy = b[0] - a[0], b[1] - a[1]
    sL = math.hypot(sdx, sdy)
    gdx, gdy = q[0] - p[0], q[1] - p[1]
    gL = math.hypot(gdx, gdy)
    if sL < 1e-9:
        return False
    if gL < 1e-9:
        return True
    cos = abs((sdx * gdx + sdy * gdy) / (sL * gL))
    return cos >= math.cos(math.radians(ang_tol))

def _hg_overlaps_material(a, b, mat_segs, L, candidates=None):
    rows = mat_segs if candidates is None else (
        mat_segs[i] for i in candidates)
    for a2, b2 in rows:
        if ((math.hypot(a[0] - a2[0], a[1] - a2[1]) < 2 and
             math.hypot(b[0] - b2[0], b[1] - b2[1]) < 2) or
            (math.hypot(a[0] - b2[0], a[1] - b2[1]) < 2 and
             math.hypot(b[0] - a2[0], b[1] - a2[1]) < 2)):
            return True
        if (abs(a[1] - b[1]) < 2 and abs(a2[1] - b2[1]) < 2 and
                abs(a[1] - a2[1]) < 2):
            x0, x1 = sorted([a[0], b[0]])
            u0, u1 = sorted([a2[0], b2[0]])
            if max(0.0, min(x1, u1) - max(x0, u0)) > L * 0.8:
                return True
        if (abs(a[0] - b[0]) < 2 and abs(a2[0] - b2[0]) < 2 and
                abs(a[0] - a2[0]) < 2):
            y0, y1 = sorted([a[1], b[1]])
            u0, u1 = sorted([a2[1], b2[1]])
            if max(0.0, min(y1, u1) - max(y0, u0)) > L * 0.8:
                return True
    return False

# flow_water.HEAD_TOUCH 과 동일. stage1↔water 순환 import 피함.
_HEAD_TOUCH = 50.0


def head_marks_of(w, spec, small_r=300.0):
    """지정 헤드 묶음의 작은 원·원호 (x, y, r) — 헤드틈 판정 재료."""
    out = []
    for hs in (spec.get("heads") or []):
        ly, c = tuple(hs["bundle"])
        for Ly, C, x, y, r in w.circles + w.arcs:
            if Ly != ly or r > small_r or r <= 0:
                continue
            if c is not None and C != c:
                continue
            out.append((float(x), float(y), float(r)))
    return out


def head_circles_of(w, spec, small_r=300.0):
    """찍은 헤드 묶음의 닫힌 원만. 접속 호는 넣지 않는다 (3F 원 vs 호)."""
    out = []
    for hs in (spec.get("heads") or []):
        b = hs.get("bundle") or ()
        if len(b) < 2:
            continue
        ly, c = b[0], b[1]
        for Ly, C, x, y, r in w.circles:
            if Ly != ly or r > small_r or r <= 0:
                continue
            if c is not None and C != c:
                continue
            out.append((float(x), float(y), float(r)))
    return out


def is_head_symbol_bar(a, b, circles, touch=_HEAD_TOUCH):
    """양 끝이 같은 닫힌 헤드 원(r+HEAD_TOUCH) 안 — 문양 «후보»일 뿐.

    확정은 head_symbol_bar_keys(끝이 허공인 토막만)가 한다. 이 함수 단독으로
    빼면 헤드로 가는 짧은 선의 안쪽 조각까지 먹는다 (3F 실측: 24~40mm 팔
    토막이 걸려 미부착 8·고립 2 → 물닿음 258→248) [2026-08-13].
    재료에서 빼지 않는다. 망 add_edge 만 건너뛴다 [2026-08-13 오너].
    상향식 밑 관통 배관은 양 끝이 원 밖이라 해당 없음.
    """
    if not circles:
        return False
    for hx, hy, hr in circles:
        lim = hr + touch
        if (math.hypot(a[0] - hx, a[1] - hy) <= lim
                and math.hypot(b[0] - hx, b[1] - hy) <= lim):
            return True
    return False


def head_bar_key(a, b):
    """문양 막대기 집합 조회용 좌표쌍 키."""
    ka = (round(a[0], 1), round(a[1], 1))
    kb = (round(b[0], 1), round(b[1], 1))
    return (ka, kb) if ka <= kb else (kb, ka)


def head_symbol_bar_keys(mat, circles, eps=2.0):
    """문양 막대기 확정 — 원 안 토막 중 짧은 선 조각이 아닌 것만.

    배관으로 남기는 두 근거 [2026-08-13 오너 — "작은 조각을 살리고
    막대기는 논리를 통해 배제한다"] · 방향(가로/세로)은 보지 않는다:
      ① 끝이 다른 재료 선에 닿는다(끝점·몸통 ≤eps) — 이어지는 조각.
      ② 끝점이 헤드 원 «둘레 위»다(|중심거리−r| ≤eps) — 도면이 원에서
         빼낸 짧은 선의 시작 토막 (3F 실측 0.0mm · 정석 ② 살리기 대상).
         옛 이음이 이 토막을 발판으로 150mm 다리를 놓아 헤드를 붙였다.
    가로막대기는 어느 쪽도 아니다 — 끝이 허공이고(드롭과는 몸통 교차),
    끝점은 둘레에서 40mm쯤 떠 있다(N3 실측 40.5 · 몸통만 원에 접함).
    """
    if not mat or not circles:
        return set()
    segs = [(a, b) for _ly, _c, a, b in mat]
    seg_grid = _bbox_grid(segs, pad=eps)
    circ_grid = _circle_grid(circles)
    max_r = max((float(hr) for _hx, _hy, hr in circles), default=0.0)
    cands = []
    for i, (_ly, _c, a, b) in enumerate(mat):
        near = tuple(_circles_near_segment(
            circ_grid, a, b, max_r + _HEAD_TOUCH))
        if is_head_symbol_bar(a, b, near):
            cands.append((i, a, b))
    out = set()
    for i, a, b in cands:
        on_ring = False
        for hx, hy, hr in _circles_near_segment(
                circ_grid, a, b, max_r + eps):
            if (abs(math.hypot(a[0] - hx, a[1] - hy) - hr) <= eps
                    or abs(math.hypot(b[0] - hx, b[1] - hy) - hr) <= eps):
                on_ring = True
                break
        if on_ring:
            continue
        free = True
        near_segs = set(_bbox_near(seg_grid, a, a, pad=eps))
        near_segs.update(_bbox_near(seg_grid, b, b, pad=eps))
        for j in near_segs:
            if j == i:
                continue
            _ly2, _c2, p, q = mat[j]
            if (seg_geom(p, q, a[0], a[1])[0] <= eps
                    or seg_geom(p, q, b[0], b[1])[0] <= eps):
                free = False
                break
        if free:
            out.add(head_bar_key(a, b))
    return out


def attach_head_symbol_bars(g, bars):
    """복사본 그래프에 문양 막대기만 붙인다. 종류 판정용. 작업 망은 그대로."""
    for a, b in bars:
        L = math.hypot(b[0] - a[0], b[1] - a[1])
        i, j = g.node(*a), g.node(*b)
        if i == j and L > 1e-9:
            px, py = g.pts[i]
            if (math.hypot(px - a[0], py - a[1])
                    <= math.hypot(px - b[0], py - b[1])):
                j = g.force_node(*b)
            else:
                i = g.force_node(*a)
        if i != j:
            g.add_edge(i, j)


def symbol_strokes(mat, eps=2.0, perp_deg=20.0):
    """재료 안에 섞인 **기호 획**을 골라낸다 [2026-08-03 오너 확정].

    오너: "기호의 작대기는 배관이 아니므로 이어서는 안 된다."
          "그건 가지배관 끝의 부속 중 하나인 **캡**이다. 도면에서 그렇게 표시한다."

    두 가지를 모양으로만 찾는다(임계 숫자 없음 — 길이 기준을 쓰지 않는다).

      ① 작대기 : 양 끝이 아무 관에도 안 붙고, 자기 **중심이 더 긴 관 위**에
                 있으며, 그 관과 **수직**인 획.
      ② 관말 캡 : 가로대 + 그 양 끝에 **같은 쪽·같은 길이** 다리 둘 +
                 가로대 **중점**에서 같은 쪽으로 나온 **다리보다 긴** 줄기
                 (줄기는 배관으로 남기고, 가로대·다리만 기호 획 [2026-08-11 오너]).

    왜 필요한가(MF2 실측): 이 획들이 배관으로 잡혀 있어, 세로 메인을 가로질러
    서로 다른 두 가지배관을 1,649mm 이어버렸다 — 오너 판정 "없는 배관을
    엉뚱한 곳에 그린 것과 같다".
    실측: MF2 작대기 1,318 + 캡 93벌(획 372) = 1,690획 · 229.6m 제외.
          MF3·3F·BF4 는 0개(오탐 없음).

    ★★ 하지 마라 [2026-08-03 실측으로 정정] — "캡이 가리키는 관 끝에서는
       잇지 마라"는 **틀린 규칙이다.** MF2 실물은 이렇게 그려져 있다:
           가지배관 1,817 ─┤ 300mm 틈(부속 자리·원호 표시) ├─ 43mm 토막 ─ 캡 ⊓
       그 300mm 는 **이어야 하는 자리**다(관이 캡까지 실제로 간다). 금지로
       구현하면 MF2 의 정당한 이음 **93곳이 전부 죽는다**(실측 확인).
       금지해야 하는 것은 '캡 쪽으로'가 아니라 **'캡 자리를 지나 반대편으로'**
       가는 이음이다. 필요해지면 5단계 부속 앉히기와 함께 넣을 것.
    반환: 기호로 판정된 mat 인덱스 집합.
    """
    segs = [(a, b) for _ly, _c, a, b in mat]
    seg_cell = 100.0
    grid = _bbox_grid(segs, cell=seg_cell, pad=eps)
    tips = defaultdict(list)
    tip_cell = 50.0
    for i, (_ly, _c, a, b) in enumerate(mat):
        for p in (a, b):
            _grid_put(tips, tip_cell, p[0], p[1], i)

    vectors = []
    for _ly, _c, a, b in mat:
        L = math.hypot(b[0] - a[0], b[1] - a[1])
        vectors.append(
            None if L < 1e-9
            else ((b[0] - a[0]) / L, (b[1] - a[1]) / L, L))

    def vec(i):
        return vectors[i]

    def unique_near(which, cell, x, y):
        """샘플 격자에 여러 번 들어간 같은 선을 첫 순서 그대로 한 번만."""
        seen = set()
        for i in _grid_near(which, cell, x, y):
            if i in seen:
                continue
            seen.add(i)
            yield i

    out = set()
    cos_perp = math.cos(math.radians(90.0 - perp_deg))
    # ① 작대기
    for i, (_ly, _c, a, b) in enumerate(mat):
        v = vec(i)
        if v is None:
            continue
        ux, uy, L = v
        if any(any(j != i and seg_geom(mat[j][2], mat[j][3], p[0], p[1])[0]
                   <= eps for j in _bbox_near(
                       grid, p, p, pad=eps, cell=seg_cell))
               for p in (a, b)):
            continue                          # 끝이 어딘가 붙었다 = 배관 후보
        mx, my = (a[0] + b[0]) / 2, (a[1] + b[1]) / 2
        for j in _bbox_near(
                grid, (mx, my), (mx, my), pad=eps, cell=seg_cell):
            if j == i:
                continue
            vj = vec(j)
            if vj is None or vj[2] < L:
                continue
            if seg_geom(mat[j][2], mat[j][3], mx, my)[0] > eps:
                continue
            if abs(ux * vj[0] + uy * vj[1]) <= cos_perp:
                out.add(i)
                break
    # ② 관말 캡
    for s in range(len(mat)):
        vs = vec(s)
        if vs is None:
            continue
        ux, uy, _Ls = vs
        A, B = mat[s][2], mat[s][3]
        legs = []
        for tip in (A, B):
            found = None
            for j in unique_near(tips, tip_cell, tip[0], tip[1]):
                if j == s:
                    continue
                vj = vec(j)
                if vj is None or abs(ux * vj[0] + uy * vj[1]) > cos_perp:
                    continue
                a2, b2 = mat[j][2], mat[j][3]
                da = math.hypot(a2[0] - tip[0], a2[1] - tip[1])
                db = math.hypot(b2[0] - tip[0], b2[1] - tip[1])
                if min(da, db) > eps:
                    continue
                far = b2 if da <= eps else a2
                side = (far[0] - tip[0]) * -uy + (far[1] - tip[1]) * ux
                if found is None or vj[2] < found[0]:
                    found = (vj[2], side, j)
            legs.append(found)
        if not all(legs) or abs(legs[0][0] - legs[1][0]) > eps or \
                legs[0][1] * legs[1][1] <= 0:
            continue
        mid = ((A[0] + B[0]) / 2, (A[1] + B[1]) / 2)
        for j in unique_near(tips, tip_cell, mid[0], mid[1]):
            if j in (s, legs[0][2], legs[1][2]):
                continue
            vj = vec(j)
            if vj is None or abs(ux * vj[0] + uy * vj[1]) > cos_perp:
                continue
            a2, b2 = mat[j][2], mat[j][3]
            da = math.hypot(a2[0] - mid[0], a2[1] - mid[1])
            db = math.hypot(b2[0] - mid[0], b2[1] - mid[1])
            if min(da, db) > eps or vj[2] <= legs[0][0] + eps:
                continue
            far = b2 if da <= eps else a2
            side = (far[0] - mid[0]) * -uy + (far[1] - mid[1]) * ux
            if side * legs[0][1] < 0:
                continue
            # 줄기(j)는 관 방향 통과 줄기 — 배관으로 남긴다.
            # 가로대(s)·다리만 기호 획 [2026-08-11 오너 핀셋].
            out.update({s, legs[0][2], legs[1][2]})
            break
    return out

def other_fire_bundles(spec):
    """기타 소방배관 묶음 [2026-08-03 오너 용어 확정].

    오너: "유저에게 '재료가 아닌데 끊는 자리'를 물으면 무슨 말인지 모른다.
          도면의 **소방배관을 전부** 고르게 하고, 그중 이번에 임포트할 계통이
          **소방대표배관**, 나머지가 **기타 소방배관**이다. 기타는 프로그램이
          알아서 '끊는 놈'으로 쓴다. 스프링클러 말고 포·물분무도 있으니
          대표 계통을 스프링클러라 부르지 마라."
    구 이름 `explainers`(설명자)도 계속 읽는다 — 옛 spec 호환.
    """
    return [tuple(e) for e in (spec.get("other_fire")
                               or spec.get("explainers") or [])]

def material_bundles_v2(w, spec, picks=None):
    """그릇 v2 — 색 전개 + 기타 소방 겸용/별도 가름 [2026-08-04·05 오너 확정].

    유저가 찍는 것은 대표 계통의 **선(들)**이고, spec 에는 그 선의
    (레이어×색)이 `material_picks` 로 기록된다(찍은 근거 보존). 재료는
    찍은 **색 전부** — 레이어 무관 · 도장 안팎 무관(§2 확정 · 3F SP×초록
    도장 1,013획 실증, 도면틀은 물 청소가 안전판). (레이어×색)을 사람이
    손으로 펴 담던 v1 `material`/`material_stamped` 나열은 폐기(대장 D15
    — "프로그램이 해야 할 일 셋을 사람이 하고 있다").

    기타 소방배관 픽(§3-C)은 프로그램이 **인식만** 한다:
      · 픽이 material_picks 와 같은 묶음 → **겸용**(apt 소화전) — 같은
        묶음이라 기계가 가를 수 없다(오너 확정). 재료 유지 · 끊는 놈 제외.
      · 색은 대표 색인데 다른 묶음 → **별도 계통** — 그 묶음을 재료에서
        빼고 끊는 놈으로 쓴다.
      · 색이 다름 → 재료 아님 · 끊는 놈(v1 그대로).
    반환: (재료 묶음 집합, 별도로 뺀 묶음 집합)
    """
    mp = {tuple(m) for m in (picks if picks is not None
                             else spec.get("material_picks") or [])}
    colors = {c for _ly, c in mp}
    sep = {of for of in other_fire_bundles(spec)
           if of[1] in colors and of not in mp}
    buckets = {(ly, c) for ly, c, _a, _b in w.segs
               if c in colors and (ly, c) not in sep}
    return buckets, sep

def stamped_material(w, spec, bundles=None):
    """재료 중 '도장 안'(블록 유래) 몫 — 명세·부검용 [그릇 v2 재정의].

    v1에서는 유저가 0-A에서 「도장 안」 줄을 고르면 그 몫만 따로 받는
    수집기였다(2026-08-03 — 3F 도장 236m 정체 미확인이 근거였다). 그 뒤
    색 전부 확정(2026-08-04)과 물 청소 실증(도면틀 607m → 잔존 5m)이
    근거를 뒤집어, v2 색 전개는 도장 안팎을 가리지 않고 통째로 받는다.
    이 함수는 이제 **받은 재료 중 블록 유래 획이 얼마인지 세는 명세
    도구**다 — D2 프로브(295mm 아닌 정체 미판정 획)와 0-B ① 문구가 쓴다.
    """
    keys = set(material_bundles_v2(w, spec)[0] if bundles is None
               else bundles)
    if not keys:
        return []
    raw_cnt = Counter((ly, c, a, b) for ly, c, a, b in w.raw_segs)
    seen = Counter()
    out = []
    for ly, c, a, b in w.segs:
        if (ly, c) not in keys:
            continue
        k = (ly, c, a, b)
        seen[k] += 1
        if seen[k] > raw_cnt.get(k, 0):      # 원시선 몫을 넘어선 것 = 도장 몫
            out.append((ly, c, a, b))
    return out

def material_with_headgap(w, spec, picks=None, report=None):
    """재료 = 색 전개(그릇 v2) + **헤드 틈으로 재료 끝에 붙는 다른 색 토막**.

    색 전개·겸용/별도 가름은 `material_bundles_v2` 가 정본이다. 0-B 시뮬이
    picks(찍은 선 목록)를 넘기면 그것을 material_picks 로 쓴다 — 화면과
    본체가 같은 전개를 돌게(판정 한 벌).

    0-B ⑥ 규칙(2026-08-02 오너: 색 비슷함 추정 없음 · 헤드 틈으로 붙은
    토막만 합친다)은 v2에서도 그대로다 — 색 전개는 **찍은 색**만 데려오므로
    다른 색 연결관 토막은 여전히 이 문으로 들어온다.

    ★승격 토막은 이름표(레이어×색)를 **재료 대표 묶음으로 바꿔 단다.**
      2단계 기호 양옆 이음이 '양쪽이 같은 묶음일 것'을 요구하기 때문이다.
      MF3 실측: 초록 483mm·2,080mm 토막의 양옆 3곳은 분기 반원이 정중앙
      (180/180mm)인데도 색이 달라 안 이어졌다 — 오너가 도면에서 짚어냈다.
    반환: (mat_raw, pick_set, mat_set, added)
      pick_set = 색 전개 결과 묶음들 · added = [{"ly","c","a","b","rep"}]
    """
    pick_set, sep = material_bundles_v2(w, spec, picks)
    mat_raw = [(ly, c, a, b) for ly, c, a, b in w.segs
               if (ly, c) in pick_set]
    if report is not None:
        # '도장 안' 몫은 명세로만 센다 — 전개가 이미 통째로 받았다.
        report["stamped"] = stamped_material(w, spec, bundles=pick_set)
        report["other_fire_sep"] = sorted(sep)
    heads = head_marks_of(w, spec)
    mat_set = set(pick_set)
    added = []
    if not heads:
        return mat_raw, pick_set, mat_set, added
    mat_segs = [(a, b) for _ly, _c, a, b in mat_raw]
    mat_ends = [(p, (ly, c)) for ly, c, a, b in mat_raw for p in (a, b)]
    mat_seg_grid = _bbox_grid(mat_segs, pad=2.0)
    mat_end_grid = _point_grid(mat_ends)
    head_grid = _circle_grid(heads)
    max_head_r = max((float(hr) for _hx, _hy, hr in heads), default=0.0)
    head_gap_pad = max(max_head_r, 30.0)
    lays = {ly for ly, _c in pick_set}
    for ly, c, a, b in w.raw_segs:
        if ly not in lays or (ly, c) in pick_set:
            continue
        L = math.hypot(b[0] - a[0], b[1] - a[1])
        near_mat = tuple(_bbox_near(mat_seg_grid, a, b, pad=2.0))
        if L < _HG_MIN_CAND_L or _hg_overlaps_material(
                a, b, mat_segs, L, candidates=near_mat):
            continue
        attached, ok, rep = 0, True, None
        for p in (a, b):
            d_best, i_best, q_best, b_best = 1e18, len(mat_ends), None, None
            for mi, (q, qb) in _grid_near(
                    mat_end_grid, _LOOKUP_CELL, p[0], p[1], rings=1):
                d = math.hypot(p[0] - q[0], p[1] - q[1])
                if (d, mi) < (d_best, i_best):
                    d_best, i_best, q_best, b_best = d, mi, q, qb
            if d_best > _HG_ATTACH_CAP:
                continue                      # 자유단
            near_heads = tuple(_circles_near_segment(
                head_grid, p, q_best, head_gap_pad))
            if not _hg_axis_align(a, b, p, q_best) or \
                    not _hg_gap_has_head(p, q_best, near_heads):
                ok = False
                break
            attached += 1
            rep = rep or b_best
        if ok and attached >= 1:
            mat_raw.append((rep[0], rep[1], a, b))
            mat_set.add((ly, c))
            added.append({"ly": ly, "c": c, "a": a, "b": b, "rep": rep})
    # ★기호 획(작대기·관말 캡)은 배관이 아니다 — 재료에서 뺀다 [오너 2026-08-03]
    sym = symbol_strokes(mat_raw)
    if sym:
        if report is not None:
            report["symbol_strokes"] = [mat_raw[i] for i in sorted(sym)]
        mat_raw = [m for i, m in enumerate(mat_raw) if i not in sym]
    elif report is not None:
        report["symbol_strokes"] = []
    return mat_raw, pick_set, mat_set, added

# ★[폐지 2026-08-02 오너 확정] expand_heads_by_size — "같은 색·같은 크기 원을
#   자동으로 헤드에 넣기". 되살리지 마라.
#   폐지 이유(MF2 실증): 설계자가 **입상배관 위치 표시**에 헤드 원을 그대로
#   갖다 썼다. 색·크기가 헤드와 완전히 같고, 입상은 관 위에 있는 게 당연하니
#   '관에 닿나'로도 못 가른다 — 기계가 가를 수 있는 축이 아예 없다.
#   대신: 같은 크기 원이 다른 묶음에 있으면 0단계 ⑦ 목록에 띄우고 **유저가
#   체크**한다(닿음 100%는 미리 체크). 자동 편입은 없다.
#   실측(자동 편입 폐지 전→후): 3F 270→270(헤드 묶음 3개를 유저가 고름) ·
#   MF3 220→220 · BF4 434→434 · MF2 390→385(입상 표시 5개 빠짐).


# ============================================================ §4-A 메인↔메인
# A1 겹침(2026-07-31 오후 재정의 — 구 '용접' 전면 교체):
#   대상 = 같은 (레이어×색) 묶음끼리만(§4-0-2). 조건 = 같은 방향 나란 포개 +
#   어긋남 0(목표 0mm — 매 런 전수 측정해 0 아닌 포개가 실존하면 그때 재결정,
#   직선 틈 10mm 확정과 같은 절차). 처리 = 하나로 합치고 합친 자리 판정 번호.
#   묶음 다른 포개 = 합치지 않음(물 안 닿는 쪽은 §5가 삭제, 둘 다 닿으면 번호).
#   구 용접의 '끝점이 남의 몸통 15mm 안' 기계 조건은 폐기(묶음 무시·T자·옆구역
#   겹침을 구별 못 해 MF 오연결 사고). T자(끝이 몸통에 닿음)는 B1 몫.


def _eb_key(i, j):
    return (min(i, j), max(i, j))

def split_keep(g, i, j, fx, fy, ebundle):
    """간선을 쪼개되 묶음을 자식 두 조각에 그대로 물려준다(§2-5)."""
    b = ebundle.pop(_eb_key(i, j), None)
    k = g.split_edge(i, j, fx, fy)
    ebundle[_eb_key(i, k)] = b
    ebundle[_eb_key(k, j)] = b
    return k

def merge_overlaps_a1(mat, knobs):
    """A1 겹침 합치기(조립 전 전처리). 입력·반환 모두 (ly, c, a, b).
    같은 묶음·같은 방향(≤1°)·어긋남 ≤a1_lat · 포개 ≥1mm인 쌍을 사슬로 묶어
    구간 합집합으로 합친다. 반환: (합친 mat, 판정 표시, 측정 보고).
    측정 보고에는 어긋남 전수 분포·0 아닌 포개(미합침)·묶음 다른 포개 자리를
    담는다 — 숫자 박제 금지(정본 §4-A1), 도면이 스스로 눈금을 준다."""
    a1_lat = knobs["a1_lat"]
    a1_scan = knobs["a1_scan"]
    n = len(mat)
    info = []
    for ly, c, a, b in mat:
        dx, dy = b[0] - a[0], b[1] - a[1]
        L = math.hypot(dx, dy)
        u = (dx / L, dy / L) if L > 1e-9 else (1.0, 0.0)
        info.append((ly, c, a, b, u, L))

    # 후보 띠는 a1_scan(기본 50mm)뿐이다. 1200mm 칸은 대형 도면에서
    # 한 칸 안 모든 선분 쌍을 만들었으므로, 판정 띠에 맞춘 작은 칸을 쓴다.
    CELL = max(100.0, float(a1_scan) * 4.0)
    grid = defaultdict(list)
    for si, (_ly, _c, a, b, _u, _L) in enumerate(info):
        x0, x1 = sorted((a[0], b[0]))
        y0, y1 = sorted((a[1], b[1]))
        for cx in range(int((x0 - a1_scan) // CELL), int((x1 + a1_scan) // CELL) + 1):
            for cy in range(int((y0 - a1_scan) // CELL), int((y1 + a1_scan) // CELL) + 1):
                grid[(cx, cy)].append(si)

    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    sin_tol = math.sin(math.radians(1.0))
    seen = set()
    offsets = []              # 같은 묶음 포개 어긋남 전수 (합침 여부 무관)
    off_rej = []              # 0 아닌 포개(미합침) — 재결정 후보
    cross_ov = []             # 묶음 다른 포개 자리
    cross_seen = set()
    for lst in grid.values():
        for ii in range(len(lst)):
            for jj in range(ii + 1, len(lst)):
                si, sj = lst[ii], lst[jj]
                key = (min(si, sj), max(si, sj))
                if key in seen:
                    continue
                seen.add(key)
                ly1, c1, ai, bi, ui, Li = info[si]
                ly2, c2, aj, bj, uj, Lj = info[sj]
                if abs(ui[0] * uj[1] - ui[1] * uj[0]) > sin_tol:
                    continue
                if Lj > Li:      # 기준 축 = 긴 쪽
                    ai, bi, ui, Li, aj, bj = aj, bj, uj, Lj, ai, bi
                nx, ny = -ui[1], ui[0]
                lat = max(abs((aj[0] - ai[0]) * nx + (aj[1] - ai[1]) * ny),
                          abs((bj[0] - ai[0]) * nx + (bj[1] - ai[1]) * ny))
                if lat > a1_scan:
                    continue
                s1 = (aj[0] - ai[0]) * ui[0] + (aj[1] - ai[1]) * ui[1]
                s2 = (bj[0] - ai[0]) * ui[0] + (bj[1] - ai[1]) * ui[1]
                lo, hi = min(s1, s2), max(s1, s2)
                ovl = min(hi, Li) - max(lo, 0.0)
                if ovl < 1.0:    # 맞닿음뿐이면 포개 아님
                    continue
                if (ly1, c1) != (ly2, c2):
                    ck = (round((ai[0] + max(lo, 0.0) * ui[0] +
                                 min(hi, Li) * ui[0]) / 2 / 1000.0),
                          round((ai[1] + max(lo, 0.0) * ui[1] +
                                 min(hi, Li) * ui[1]) / 2 / 1000.0))
                    if ck not in cross_seen:
                        cross_seen.add(ck)
                        mx = ai[0] + (max(lo, 0.0) + min(hi, Li)) / 2 * ui[0]
                        my = ai[1] + (max(lo, 0.0) + min(hi, Li)) / 2 * ui[1]
                        cross_ov.append({"at": (round(mx, 1), round(my, 1)),
                                         "bundles": [(ly1, c1), (ly2, c2)],
                                         "lat": round(lat, 2)})
                    continue
                offsets.append(round(lat, 3))
                if lat > a1_lat:
                    off_rej.append({"lat": round(lat, 2),
                                    "at": (round((ai[0] + bi[0]) / 2, 1),
                                           round((ai[1] + bi[1]) / 2, 1))})
                    continue
                ri, rj = find(si), find(sj)
                if ri != rj:
                    parent[ri] = rj

    groups = defaultdict(list)
    for si in range(n):
        groups[find(si)].append(si)
    out, marks = [], []
    n_groups = 0
    for members in groups.values():
        if len(members) == 1:
            ly, c, a, b, _u, _L = info[members[0]]
            out.append((ly, c, a, b))
            continue
        n_groups += 1
        base = max(members, key=lambda s: info[s][5])
        ly, c, a0, _b0, u, _L = info[base]
        nx, ny = -u[1], u[0]
        wsum = csum = 0.0
        iv = []
        for s in members:
            _l2, _c2, a, b, _u2, L2 = info[s]
            mc = (((a[0] + b[0]) / 2 - a0[0]) * nx + ((a[1] + b[1]) / 2 - a0[1]) * ny)
            wsum += L2
            csum += mc * L2
            s1 = (a[0] - a0[0]) * u[0] + (a[1] - a0[1]) * u[1]
            s2 = (b[0] - a0[0]) * u[0] + (b[1] - a0[1]) * u[1]
            iv.append((min(s1, s2), max(s1, s2)))
        cbar = csum / wsum if wsum > 1e-9 else 0.0
        iv.sort()
        cur_lo, cur_hi = iv[0]
        spans = []
        for lo, hi in iv[1:]:
            if lo <= cur_hi + 1.0:
                cur_hi = max(cur_hi, hi)
            else:
                spans.append((cur_lo, cur_hi))
                cur_lo, cur_hi = lo, hi
        spans.append((cur_lo, cur_hi))
        for lo, hi in spans:
            pa = (a0[0] + lo * u[0] + cbar * nx, a0[1] + lo * u[1] + cbar * ny)
            pb = (a0[0] + hi * u[0] + cbar * nx, a0[1] + hi * u[1] + cbar * ny)
            out.append((ly, c, pa, pb))
            marks.append(((pa[0] + pb[0]) / 2, (pa[1] + pb[1]) / 2))
    report = {"offsets": offsets, "off_rej": off_rej, "cross_overlaps": cross_ov,
              "n_groups": n_groups}
    return out, marks, report

def touch_tee_b1(g, mat_raw, knobs, thgrid, ebundle):
    """B1 T자 — 끝이 몸통에 닿는 경우(2026-07-31 오후: A1에서 분리, B1로).
    닿음 자체가 그려진 증거라 기호 없이 접속(맞닿은 꺾임 자연 접속과 같음).

    ★색 장벽 폐지 [2026-08-05 오너 판정 — "다른색이 있어도 이을 수 있다.
      규칙 수정 필요함"]. 구 §4-0-2 는 닿음에도 같은 묶음만 허용해서, 색만
      다른 T자 710곳 중 326곳(46%)을 못 봤다(apt 53%·MF 81% — 실측 2026-08-05).
      D11(3단계가 그 갈라진 끝을 오인해 없는 관 생성)의 뿌리였다.
      §4-0-2 는 A2 직선 틈(떨어진 끝끼리)에는 그대로 남는다 — 여기만 예외다:
      끝점이 몸통 위 1mm 안 = 설계자가 거기서 갈라진다고 그린 것.

    명중 너머로 제 축 원시 선이 이어지면 통과(관말 탭 아님)라 잇지 않는다.
    ★단 공선(같은 방향) 닿음은 너머 검사를 건너뛴다 [2026-08-05] — 색만 다른
      한 관이 한 줄로 겹쳐 그려진 경우 '너머'가 바로 그 몸통 자신이라 전부
      통과로 오판된다(apt 같은방향 닿음 359곳 실측). 트림돼 떨어진
      T(30~450)는 join_round1 T 단계 몫(기호 필요)."""
    touch = knobs["b1_touch"]
    ends = []                     # (P, 축방향 u, 묶음)
    for ly, c, a, b in mat_raw:
        dx, dy = b[0] - a[0], b[1] - a[1]
        L = math.hypot(dx, dy)
        if L < 1e-9:
            continue
        ends.append((b, (dx / L, dy / L), (ly, c)))
        ends.append((a, (-dx / L, -dy / L), (ly, c)))
    # 그래프 간선 격자(쪼갤 대상 검색)
    egrid = defaultdict(list)

    def put_edge(ea, eb):
        a2, b2 = g.pts[ea], g.pts[eb]
        x0, x1 = min(a2[0], b2[0]) - touch, max(a2[0], b2[0]) + touch
        y0, y1 = min(a2[1], b2[1]) - touch, max(a2[1], b2[1]) + touch
        for gx in range(int(x0 // _B1_CELL), int(x1 // _B1_CELL) + 1):
            for gy in range(int(y0 // _B1_CELL), int(y1 // _B1_CELL) + 1):
                egrid[(gx, gy)].append((ea, eb))

    for (i, j) in g.edges:
        put_edge(i, j)
    done = set()
    n_touch = 0
    side = []
    for P, (ux, uy), bnd in ends:
        best = (1e18, None)
        cell_rows = egrid.get(
            (int(P[0] // _B1_CELL), int(P[1] // _B1_CELL)), ())
        for (i, j) in set(cell_rows):
            if _eb_key(i, j) not in g.edges:
                continue
            # ★묶음 관문 폐지 [2026-08-05 오너 판정] — 닿음은 색이 달라도 잇는다.
            d, u2, _f = seg_geom(g.pts[i], g.pts[j], P[0], P[1])
            if d <= touch and 0.01 <= u2 <= 0.99 and d < best[0]:
                if math.hypot(P[0] - g.pts[i][0], P[1] - g.pts[i][1]) <= SNAP:
                    continue
                if math.hypot(P[0] - g.pts[j][0], P[1] - g.pts[j][1]) <= SNAP:
                    continue
                best = (d, (i, j))
        if best[1] is None:
            continue
        i, j = best[1]
        # 명중 너머 연속 검사 — 이어지면 통과(잇지 않음).
        # ★공선 닿음(맞은 몸통과 같은 방향 ≤20°)은 검사를 건너뛴다 — '너머'가
        #   그 몸통 자신이라 헛돈다(각도 눈금은 _cov_axis 평행 판정과 동일).
        a_col = ang_between((ux, uy), (g.pts[j][0] - g.pts[i][0],
                                       g.pts[j][1] - g.pts[i][1]))
        if 20.0 < a_col < 160.0:
            if any(_cov_axis(thgrid, P[0] + ux * t0, P[1] + uy * t0, ux, uy)
                   for t0 in (90.0, 200.0, 315.0)):
                continue
        ck = (round(P[0] / 50.0), round(P[1] / 50.0))
        if ck in done:
            continue
        done.add(ck)
        k = split_keep(g, i, j, P[0], P[1], ebundle)
        np_ = g.node(P[0], P[1])
        g.add_edge(np_, k)
        ebundle[_eb_key(np_, k)] = bnd
        n_touch += 1
        side.append({"at": (round(P[0], 1), round(P[1], 1)),
                     "bundle": list(bnd), "d": round(best[0], 1)})
        # 쪼갠 조각을 격자에 등록 — 뒤 닿음 후보가 볼 수 있게
        for (ea, eb) in ((i, k), (k, j)):
            put_edge(ea, eb)
    return n_touch, side

def _cov_axis(thgrid, x, y, ax, ay):
    """점(x,y)이 축(ax,ay) 방향 원시 선으로 덮여 있나(평행 ≤20°·거리 ≤30)."""
    rows = thgrid.get((int(x // _B1_CELL), int(y // _B1_CELL)), ())
    for (a3, b3, _bnd) in set(rows):
        vx0, vy0 = b3[0] - a3[0], b3[1] - a3[1]
        Lv = math.hypot(vx0, vy0) or 1.0
        if abs(vx0 * ax + vy0 * ay) / Lv < math.cos(math.radians(20.0)):
            continue
        if seg_geom(a3, b3, x, y)[0] <= 30.0:
            return True
    return False

# ============================================== 단계 사진 [2026-08-05 오너 확정]
# "결과 판정용 그림이 본체 로직과 다른 사고가 계속" → 별도 사다리 스크립트
# (_tmp_stage1~4) 폐지. 같은 재구현 어긋남 사고가 세 번 반복됐다:
#   ① 크기확장이 단계에 없어 3세션 오진  ② 헤드틈합침이 본체에 없어 재료 다름
#   ③ B1 T닿음이 단계에 없어 apt 덩어리 223 vs 본체 108
# 그림은 이제 본체가 도는 중간에 스스로 찍는다 — 어긋날 상대가 없다.


def graph_comps(g):
    """그래프의 덩어리(연결 성분) — 큰 것부터, 간선 목록으로."""
    par = {}

    def find(x):
        par.setdefault(x, x)
        while par[x] != x:
            par[x] = par[par[x]]
            x = par[x]
        return x

    for i, j in g.edges:
        ri, rj = find(i), find(j)
        if ri != rj:
            par[ri] = rj
    bags = defaultdict(list)
    for i, j in g.edges:
        bags[find(i)].append((i, j))
    return sorted(bags.values(), key=lambda es: -sum(
        math.hypot(g.pts[j][0] - g.pts[i][0], g.pts[j][1] - g.pts[i][1])
        for i, j in es))

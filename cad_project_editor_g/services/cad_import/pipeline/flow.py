# 물길 주도 임포트 — 시제품 2호 [2026-08-07 오너 설계]
# ============================================================================
# 순서가 본체와 정반대다.
#   본체    : 도면 전체를 미리 잇고 → 그 다음 물을 흘린다.
#   여기    : 물을 먼저 흘리고 → 물이 «닿은 끝점»에서만 판정한다.
#
# ★단계 梯子 0~6 [2026-08-08 오너 확정 · 이름·성적만 — 이음 로직 불변]
# ---------------------------------------------------------------------------
#   0 찍기
#   1 날것으로 펼치기          `pipeline.expand.stage1_body`
#   1-1 헤드 종류 분류         `stage11_classify_heads` (classify 재사용·미지정 알림)
#   2 접속부에서 배관 연결      spots · thru · join_all (헤드 원 통과 이음 안 함)
#   3 헤드 접속                 관말 등록·확인 · 헤드 경유 이음 금지 (이음3=0)
#   4 끊긴 배관 잇기            `stage4_body` ← 옛 stage3 / 옛 이음3
#   5 상향식 헤드 접속          `stage5_body` ← 옛 stage4 / 옛 이음4
#   6 물길 확인                 bodies · flow · render_now · `_tmp_score`
#
# 오너 규칙 (2026-08-07 확정)
# ---------------------------------------------------------------------------
#   물이 끝점에 닿으면 그 자리에 «접속표시»가 있나 본다.
#     · 있으면  → 선이 안 붙어 있어도 그냥 통과. 그 표시에 모인 팔 전부로 분기.
#                 (표시가 곧 허가다 — 반경·방향을 증명하지 않는다)
#     · 없으면  → 프로그램이 추측하지 않는다. 유저에게 «어디와 이을까» 묻는다.
#   헤드는 관말이다 — 헤드 원을 티처럼 통과시켜 좌우를 잇지 않는다.
#   끝 = 모든 헤드에 물이 닿았을 때. 「덩어리 1개」가 목표가 아니다.
#
# 새로 만드는 것이 없다 [오너 지시]
# ---------------------------------------------------------------------------
#   1단계   : `pipeline.expand.stage1_body` — pipeline.stage1 소유
#   접속표시: `pipeline.heads`
#   4·5단계 : `pipeline.stage45`
#   ★단 하나 보탠 것 = 「재료 레이어 CIRCLE」. 본체는 「CIRCLE은 재료원이
#     아니다」로 빼지만, 오너 규칙에서는 그려진 표시면 허가다. MF2 급수원
#     입상관 원(r300 · "A"상세 참조)이 바로 이것이라 빼면 물이 출발조차 못 한다.
# 사이드카·도장·본체 그림을 **쓰지 않는다.**

# 성적·로그·그림에 찍는 단계 이름 (동작과 무관 — 라벨만)
STAGE_NAME = {
    0: "찍기",
    1: "날것으로 펼치기",
    "1-1": "헤드 종류 분류",  # 1과 2 사이 · 2~6 번호 유지 [2026-08-09]
    2: "접속부에서 배관 연결",
    3: "헤드 접속",
    4: "끊긴 배관 잇기",
    5: "상향식 헤드 접속",
    6: "물길 확인",
}

import copy
import json
import math
import os
import sys
from collections import Counter, defaultdict, deque

from services.cad_import.pipeline import heads
from services.cad_import.pipeline import stage1 as s1
from services.cad_import.pipeline import stage45 as s45
from services.cad_import.pipeline.expand import gput, gnear, seg_dist
from services.cad_import.pipeline.handoff import default_edits_dir
from services.cad_import.kinds import (
    disk_key as _disk_key,
    normalize_head_kind,
    normalize_head_slot,
    require_head_kinds,
)

DWG = s1.DWG_DIR

ARM_CTR = 5.0       # 팔이 «표시 중심»에 앉았다고 볼 거리
ARM_RIM = 12.0      # 팔이 «표시 테두리»에 앉았다고 볼 오차
ARM_TIE = 1.0       # 팔 후보 «박빙» — 당선·차선 거리 차가 이 미만일 때만 알린다 [오너 2026-08-17]
HEAD_RIM = 50.0     # ★헤드 테두리만 따로 — 실측으로 정한다(spot_arms 주석)
SRC_SNAP = 2500.0   # 급수원 스냅 한계 — 본체와 같은 값
ASK_FAR = 40000.0   # 질문거리에서 「이을 만한 상대」를 찾는 한계

# 오너가 CAD 화면에 흰선으로 찍어 준 급수원 [2026-08-07]
#   MF2 = 입상관 원("A"상세 참조 · 71053,64403 r300) 아래 703mm 관.
#   화살표는 동쪽(100A 본관) 쪽을 가리켰다.
SOURCES = {
    "MF2": [("흰선(입상관 아래 703mm)", 71053, 63752)],
}


# --------------------------------------------------------- 접속표시 명단
def spots_body(st, cap=None, owner=True, outside=False):
    """접속표시 명단.

    outside=False — ★원칙 1 「찍은 것만 쓴다」 [2026-08-07 오너 확정]
      오너가 찍은 묶음 밖의 도형은 **쓰지 않는다.** 필요해 보이면 조용히
      주워 오지 말고 `unpicked_hints()` 로 «여기 접속부를 안 찍으셨습니다»
      목록을 낸다. 오너: "찍은 것 밖은 안 본다. 만일 내가 실수했다면
      얘기는 해 줘야지."
      outside=True 는 옛 방식(밖까지 긁어 오기) — 전후 대조에만 쓴다.

    owner=True — ★오너 안 [2026-08-07 확정]
    ---------------------------------------------------------------------
    오너: "헤드를 찍으면 헤드 원만 딱 집어서 가져오면 되지 않나? 헤드를
           제외한 나머지는 전부 배관에 붙이면 배관 규칙에 의해서 접속기호와
           헤드랑 연결된 배관까지 규칙을 따라 연결하고, 마지막에 헤드는
           연결배관선이 헤드에 붙어 있으니 그것을 접속점으로 본다."

      · 헤드     = 찍은 반지름 자를 통과한 **닫힌 원(CIRCLE)** 각각. 그것뿐.
      · 접속표시 = 나머지 작은 도형 전부(호·원·짧은 획). 레이어를 안 본다.

    3F는 헤드와 접속부가 레이어(SP)·색(3)·반지름(150)까지 똑같고 **원이냐
    호냐**만 다르다. 도형 종류로 가르면 정확히 나뉜다. 옛 방식은 레이어로
    갈라서 접속부 반호 353개가 통째로 「헤드」 이름표를 달았고, 본체 4단계가
    그것을 헤드로 오독해 3F 오이음 150곳·MF3 99곳을 냈다.

    잡동사니가 새어 드는 걱정은 실측으로 이미 답이 있다 — 재료 밖 작은 호를
    MF3 +2,674개·BF4 +2,792개 들이부어도 **이음은 0곳 늘었다**(2026-08-03).
    접속표시로 인정받으려면 그 자리에 팔이 둘 이상 모여야 하는데 건축 문짝
    호·치수선 호는 배관 틈에 앉아 있지 않아 아무것도 못 잇는다.

    owner=False — 옛 방식. 전후 대조에만 쓴다.

    cap = 「작은 원호」 반지름 문턱. None 이면 본체 knobs 값(300mm).
    """
    w, spec, knobs = st["w"], st["spec"], st["knobs"]
    small_r = float(cap or knobs["small_r"])
    small_len = knobs["small_len"]
    clusters0 = heads.collect_head_clusters(w, spec, knobs)
    head_cls, marks, _hinfo = heads.split_head_circles(clusters0, knobs)
    clusters = head_cls + marks
    # ★`spec.get` 이다 — 헤드를 하나도 안 찍으면 찍기가 이 칸을 **아예 안 쓴다**
    #   (`pick/board.spec()` 의 `if self.heads:`). 종전에는 여기서
    #   `KeyError: 'heads'` 로 죽었고, 사람은 「자동 인식이 0 개인 도면」에서
    #   사유 대신 그 낱말만 봤다(BLOCKED §16). 250 행은 이미 이렇게 읽고 있었다.
    head_lays = {tuple(hs["bundle"])[0]
                 for hs in (spec.get("heads") or ()) if "r" in hs}
    mat_set = set(st["mat_bundles"])
    mat_layers = {ly for ly, _c in mat_set}

    S = []

    def put(kind, cx, cy, r, sa=None, sweep=None):
        rec = dict(k=kind, cx=float(cx), cy=float(cy), r=float(r))
        if sa is not None and sweep is not None:
            rec["sa"] = float(sa)
            rec["sweep"] = float(sweep)
        S.append(rec)

    if owner:
        hcov = heads.head_cover_disks(head_cls, small_r)
        hkey = {(round(x, 1), round(y, 1), round(r, 1)) for x, y, r in hcov}
        angs = getattr(w, "arc_ang", ())
        for i, (ly, _c, cx, cy, r) in enumerate(w.arcs):
            if not 0 < r <= small_r:
                continue
            ang = angs[i] if i < len(angs) else None
            sa = sw = None
            if ang:
                sa, sw = ang[0], ang[1]
            if ly in mat_layers:
                put("호", cx, cy, r, sa, sw)
            elif outside:
                put("밖호", cx, cy, r)
        for ly, _c, cx, cy, r in w.circles:
            if not 0 < r <= small_r:
                continue
            if (round(cx, 1), round(cy, 1), round(r, 1)) in hkey:
                continue          # ★헤드 원만 뺀다 — 나머지 원은 배관 세계로
            if ly in mat_layers:
                put("원", cx, cy, r)
            elif outside:
                put("밖원", cx, cy, r)
        # ★기호 획 = 본체가 «모양으로» 골라 재료에서 뺀 것만 쓴다.
        #   짧은 선을 시제품이 제 눈으로 다시 고르면 «짧은 배관»까지 접속표시가
        #   되어 없는 관을 그린다 — MF2 배관 2,407→3,211m(+804m) 실측.
        for _ly, _c, a, b in st.get("sym_strokes") or ():
            ln = math.hypot(b[0] - a[0], b[1] - a[1])
            if ln > 0:
                put("획", (a[0] + b[0]) / 2, (a[1] + b[1]) / 2, ln / 2)
                # 획 방향(단위벡터) — thru_arms 규칙 A 스침 판정에만 쓴다
                S[-1]["dx"] = (b[0] - a[0]) / ln
                S[-1]["dy"] = (b[1] - a[1]) / ln
        if outside:
            for ly, c, a, b in w.segs:
                if (ly, c) in mat_set:
                    continue
                ln = math.hypot(b[0] - a[0], b[1] - a[1])
                if 0 < ln <= small_len:
                    put("밖획", (a[0] + b[0]) / 2, (a[1] + b[1]) / 2, ln / 2)
        for x, y, r in hcov:
            put("헤드", x, y, r)
        return S, hcov

    for _ly, _c, cx, cy, r in heads.iter_material_symbol_arcs(
            w, mat_set, head_lays, small_r):
        put("재료원호", cx, cy, r)
    for _ly, _c, cx, cy, r in heads.iter_outside_symbol_arcs(
            w, mat_set, head_lays, small_r):
        put("밖원호", cx, cy, r)
    for ly, _c, cx, cy, r in w.circles:          # ★오너 규칙으로 보탠 것
        if ly in head_lays or ly not in mat_layers or not 0 < r <= small_r:
            continue
        put("재료원판", cx, cy, r)
    for ly, _c, cx, cy, r in w.circles + w.arcs:
        if ly in head_lays or ly in mat_layers or not 0 < r <= small_r:
            continue
        put("바깥원", cx, cy, r)
    for ly, c, a, b in w.segs:
        if (ly, c) in mat_set or ly in head_lays:
            continue
        ln = math.hypot(b[0] - a[0], b[1] - a[1])
        if 0 < ln <= small_len:
            put("획", (a[0] + b[0]) / 2, (a[1] + b[1]) / 2, ln / 2)
    for x, y, r in heads.head_disks(clusters, small_r):
        put("헤드", x, y, r)

    hcov = heads.head_cover_disks(head_cls, small_r)   # 헤드 실체 원 = 물 닿음 판정
    return S, hcov


def ho_from_spots(spots):
    """찍힌 호 위치·각. payload['ho'] (도면 mm). 변환이 DXF를 다시 안 연다."""
    out = []
    for s in spots or ():
        if s.get("k") != "호":
            continue
        rec = {"cx": float(s["cx"]), "cy": float(s["cy"]), "r": float(s["r"])}
        if s.get("sa") is not None and s.get("sweep") is not None:
            rec["sa"] = float(s["sa"])
            rec["sweep"] = float(s["sweep"])
        out.append(rec)
    return out


def ho_from_spec(spec):
    """찍기 v2 의 실제 각만. 옛 파일·각 없는 항은 비운다 (추정 안 함)."""
    if not spec or spec.get("format") != "v2":
        return []
    return [h for h in ho_from_spots(
        dict(r, k="호") for r in (spec.get("ho") or ()))
        if h.get("sa") is not None and h.get("sweep") is not None]


def unpicked_hints(st):
    """★「여기 접속부를 안 찍으셨습니다」 [오너 2026-08-07 · 원칙 1].

    찍은 묶음 «밖»의 작은 도형 가운데, 그 자리에 배관 끝이 둘 이상 모여
    있는 것 — 즉 찍혀 있었다면 **이었을** 자리를 찾는다.

    프로그램은 이것을 **쓰지 않는다.** 조용히 주워 와서 메우면 유저는 자기가
    무엇을 빠뜨렸는지 영영 모른다. 목록으로 내서 오너가 판정하시게 한다.

    반환: [{bundle, kind, n, join, pts}] — 묶음별로 «찍으면 몇 곳이 이어지나».
    """
    w, knobs = st["w"], st["knobs"]
    small_r, small_len = knobs["small_r"], knobs["small_len"]
    mat_set = set(st["mat_bundles"])
    mat_layers = {ly for ly, _c in mat_set}
    # ★헤드로 찍으신 것도 «찍은 것»이다 — 빠뜨림으로 올리면 안 된다.
    head_lays = {tuple(hs["bundle"])[0] for hs in (st["spec"]["heads"] or ())}
    skip_lay = mat_layers | head_lays

    out = []

    def put(ly, c, kind, cx, cy, r):
        out.append(dict(k=kind, b=(ly, c), cx=float(cx), cy=float(cy),
                        r=float(r)))

    for ly, c, cx, cy, r in w.arcs:
        if 0 < r <= small_r and ly not in skip_lay:
            put(ly, c, "호", cx, cy, r)
    for ly, c, cx, cy, r in w.circles:
        if 0 < r <= small_r and ly not in skip_lay:
            put(ly, c, "원", cx, cy, r)
    for ly, c, a, b in w.segs:
        ln = math.hypot(b[0] - a[0], b[1] - a[1])
        if 0 < ln <= small_len and (ly, c) not in mat_set \
                and ly not in head_lays:
            put(ly, c, "획", (a[0] + b[0]) / 2, (a[1] + b[1]) / 2, ln / 2)

    arms, _ns = spot_arms(st["g"], out)
    bag = defaultdict(lambda: dict(n=0, join=0, pts=[]))
    for si, sp in enumerate(out):
        row = bag[(sp["b"], sp["k"])]
        row["n"] += 1
        if len(arms[si]) >= 2:
            row["join"] += 1
            if len(row["pts"]) < 8:
                row["pts"].append((sp["cx"], sp["cy"]))
    rows = [dict(bundle=k[0], kind=k[1], **v) for k, v in bag.items()
            if v["join"]]
    rows.sort(key=lambda r: -r["join"])
    return rows


def spot_arms(g, spots, rim=None, ctr=None, head_rim=None):
    """표시마다 «모인 팔» — 관 끝이 표시 중심(0) 또는 테두리(r)에 앉은 노드.

    rim = 테두리 여유. 오너 규칙: "끊어져 있어도 아주 작은 거리라 그정도는
    여유로 주면 된다". 실측으로 정한다.

    head_rim = ★헤드 테두리 여유는 따로 둔다 [2026-08-07].
      오너: "유저가 헤드에 선을 딱 붙이지 않는 경우도 있어. 그런 경우에도
             육안으로는 거의 붙은 것처럼 그리기 때문에, 헤드 원호에서 아주
             근접해서 떨어진 것은 그냥 붙이면 되겠지."
      헤드에 관이 둘 이상 붙으면 그 헤드는 접속점이다 — 헤드가 메인배관 위에
      걸쳐 그려져 관이 끊긴 자리(헤드겹침)가 여기서 이어진다.
    """
    rim = ARM_RIM if rim is None else float(rim)
    ctr = ARM_CTR if ctr is None else float(ctr)
    hrim = HEAD_RIM if head_rim is None else float(head_rim)
    cell = 400.0
    ng = defaultdict(list)
    for i, (x, y) in enumerate(g.pts):
        gput(ng, cell, x, y, i)
    arms, node_spots = [], defaultdict(list)
    for si, sp in enumerate(spots):
        rm = hrim if sp["k"] == "헤드" else rim
        rings = 1 + int((sp["r"] + rm) // cell)
        best = []
        for n in set(gnear(ng, cell, sp["cx"], sp["cy"], rings=rings)):
            d = math.hypot(g.pts[n][0] - sp["cx"], g.pts[n][1] - sp["cy"])
            if d <= ctr or abs(d - sp["r"]) <= rm:
                best.append(n)
        arms.append(best)
        for n in best:
            node_spots[n].append(si)
    return arms, node_spots


THRU_OFF = 25.0     # 「지나간다」고 볼 중심 벗어남
THRU_IN = 60.0      # 관 끝에서 이만큼 안쪽을 지나야 «지나가는 관»
THRU_STROKE_D = 5.0  # 규칙 A — 획 중심에서 이 거리 밖이면 스침 후보
THRU_STROKE_COS = math.cos(math.radians(15.0))  # 규칙 A — 평행 판정 (각도차 ≤ 15°)


def thru_arms(pts, edges, spots, arms, off=THRU_OFF, inside=THRU_IN):
    """★표시 «안을 지나가는 관»도 팔이다 [오너 2026-08-07].

    오너: "메인과 가지관이 만나는 접속부인데 왜 잇지 않은건가?
           접속부는 무조건 이어라고 했잖아."

    MF2 실측으로 드러난 구멍이다. 보라 메인 y=80789.5 줄은 300mm 틈 24곳
    가운데 22곳이 메인-메인으로 이어졌는데, 그 틈마다 초록 가지배관이
    **끝점 없이 곧게 지나간다.** 팔을 «관 끝»으로만 모으면 지나가는 관은
    팔이 0개라 메인과 영원히 안 만난다 — 도면에는 십자로 그려져 있는데도.

    그래서 표시 중심을 지나는 관을 그 자리에서 쪼개어 새 노드를 팔로 넣는다.
    교차 관행(그냥 스치는 관)과 헷갈리지 않게 두 문을 좁게 잠근다.
      · 중심에서 벗어남 ≤ off      (표시 정중앙을 지나야 한다)
      · 관 끝에서 inside 보다 안쪽 (끝이 앉은 것은 이미 팔이다)

    ★짧은 관(L < 2·inside)은 쪼갤 «안쪽»이 없다 [2026-08-07 밤 · 3F 실측].
      그래도 중심을 스치며 헤드로 가는 세로(예: L100, 중심과 14mm)는 팔이다.
      쪼개지 않고 **중심에 더 가까운 끝**을 팔로 넣는다 → `join_all` 이
      가로를 이은 뒤 허브에 붙이는 것과 같은 길이다. 헤드는 붙어야 한다.
    반환: (새 pts, 새 edges, 새 arms, 쪼갠 수)
    """
    cell = 1000.0
    eg = defaultdict(list)
    for (i, j) in edges:
        a, b = pts[i], pts[j]
        n = max(1, int(math.hypot(b[0] - a[0], b[1] - a[1]) / cell))
        for k in range(n + 1):
            t = k / n
            gput(eg, cell, a[0] + (b[0] - a[0]) * t,
                    a[1] + (b[1] - a[1]) * t, (i, j))

    want = defaultdict(list)          # 변 → [(t, 표시번호)]
    short_arm = defaultdict(list)     # 표시 → 짧은 관의 가까운 끝
    for si, sp in enumerate(spots):
        if sp["k"] == "헤드":
            continue                  # ★헤드는 관을 쪼개지 않는다 — 관말이다
        cx, cy = sp["cx"], sp["cy"]
        for (i, j) in set(gnear(eg, cell, cx, cy, rings=1)):
            ax, ay = pts[i]
            bx, by = pts[j]
            L = math.hypot(bx - ax, by - ay)
            if L < 1e-9:
                continue
            t = ((cx - ax) * (bx - ax) + (cy - ay) * (by - ay)) / (L * L)
            if t < 0.0 or t > 1.0:
                continue              # 선분이 중심 옆을 안 지남
            px, py = ax + (bx - ax) * t, ay + (by - ay) * t
            d = math.hypot(px - cx, py - cy)
            if d > off:
                continue
            # ★규칙 A [2026-08-17 오너] — 획과 평행하게 스치는 남의 관은 팔이
            #   아니다. 획 spot · 평행(각도차 ≤15°) · 중심에서 5mm 밖이면 건너뛴다.
            #   MF4 사고: 획 중심 22.8mm 밖 헤드 팔을 잡아 100mm 메인에 이음.
            if sp["k"] == "획" and sp.get("dx") is not None \
                    and d > THRU_STROKE_D:
                if abs(sp["dx"] * (bx - ax) + sp["dy"] * (by - ay)) / L \
                        >= THRU_STROKE_COS:
                    continue
            if L < 2 * inside:
                # 짧은 관 — 쪼개지 않고 가까운 끝을 팔로 (허브 붙임 재료)
                di = math.hypot(ax - cx, ay - cy)
                dj = math.hypot(bx - cx, by - cy)
                short_arm[si].append(i if di <= dj else j)
                continue
            s = t * L
            if s < inside or s > L - inside:
                continue              # 끝에 앉은 것 = 이미 팔이다
            want[(i, j)].append((t, si))

    pts2 = list(pts)
    arms2 = [list(a) for a in arms]
    edges2 = set(edges)
    made = 0
    for (i, j), lst in want.items():
        edges2.discard((i, j))
        edges2.discard((j, i))
        ax, ay = pts[i]
        bx, by = pts[j]
        L = math.hypot(bx - ax, by - ay)
        lst.sort()
        prev, k = i, 0
        while k < len(lst):
            t = lst[k][0]
            same = []                 # 같은 자리(반원 둘 등)는 노드 하나로
            while k < len(lst) and (lst[k][0] - t) * L <= 1.0:
                same.append(lst[k][1])
                k += 1
            pts2.append((ax + (bx - ax) * t, ay + (by - ay) * t))
            new = len(pts2) - 1
            edges2.add((prev, new))
            for si in same:
                arms2[si].append(new)
            prev = new
            made += 1
        edges2.add((prev, j))
    for si, ends in short_arm.items():
        have = set(arms2[si])
        for n in ends:
            if n not in have:
                arms2[si].append(n)
                have.add(n)
    return pts2, frozenset(edges2), arms2, made


# --------------------------------------------------------- 물길
def snap(g, sx, sy):
    eg = defaultdict(list)
    for (i, j) in g.edges:
        a, b = g.pts[i], g.pts[j]
        steps = max(1, int(math.hypot(b[0] - a[0], b[1] - a[1]) / 1000.0))
        for k in range(steps + 1):
            t = k / steps
            gput(eg, 1000.0, a[0] + (b[0] - a[0]) * t,
                    a[1] + (b[1] - a[1]) * t, (i, j))
    best = (1e18, None)
    for (i, j) in set(gnear(eg, 1000.0, sx, sy, rings=3)):
        d, _t = seg_dist(g.pts[i], g.pts[j], sx, sy)
        if d < best[0]:
            best = (d, (i, j))
    return best


def mlen(pts, edges):
    return sum(math.hypot(pts[j][0] - pts[i][0], pts[j][1] - pts[i][1])
               for i, j in edges) / 1000.0


HEAD_TOUCH = 50.0   # 헤드 «붙었다» 자 [오너 2026-08-07]
# 오너: "내 눈에는 헤드에 배관이 붙어 있다. 유저가 헤드에 선을 딱 붙이지
#        않는 경우에도 육안으로 거의 붙은 것처럼 그리기 때문에, 헤드 원호에서
#        아주 근접해서 떨어진 것은 그냥 붙이면 되겠지."
# 7도면 실측(_tmp_owner_split.py): 관 끝은 헤드 «테두리 0.0mm» 또는 «헤드
# 중심»에 앉는다. 테두리에서 떨어진 거리의 99%가 —
#   3F 0.0 · BF4 0.0 · MF4 9.7 · MF3 51.8 · apt 65.0 · MF2 94.8mm.
# 옛 값 300mm 는 헤드 반지름의 두 배라 옆 가지관까지 「이 헤드 젖었다」로
# 셌다. 그래서 헤드 셈 합이 실제보다 컸다(MF2 489 vs 399).


def head_nodes(pts, hcov, tol=None, edges=None, upright=()):
    """헤드마다 «그 헤드에 물을 대는 노드» — 물 닿음을 재는 자.

    ★잣대를 `spot_arms` 와 같게 맞춘다 [2026-08-07]. 관 끝은 헤드 «중심»이나
    «테두리»에 앉지 그 사이에 뜨지 않는다(7도면 실측). 옛 방식은 중심에서
    반지름+300mm 안을 **전부** 긁어서, 헤드 옆을 스쳐 가는 남의 가지관까지
    「이 헤드 젖었다」로 셌다 — MF2 헤드 셈 합 489 vs 실제 399.

    ★★상향식 헤드에는 «관 끝»이 없다 [오너 확정 2026-08-07 밤].
      배관 위에 티로 올라앉고 거기서 z축으로 갈라져 오르므로, 평면도에서는
      배관이 헤드 밑을 그냥 지나간다. 그 자리에 끝점이 없다고 「물이 안
      닿았다」로 세면 멀쩡히 젖은 헤드를 마른 것으로 센다.
      그래서 상향식일 수 있는 헤드에 한해 «원 안을 지나는 관»도 물길로 본다.
      하향식에는 이 문을 열지 않는다 — 우연히 스쳐 가는 남의 관을 그 헤드에
      달아 버리기 때문이다.

    ★★★중심 접속 확정 [2026-08-13 오너]. 헤드 «중심»(≤ARM_CTR)에 간선 달린
      노드가 있으면 그 헤드는 그 노드에서만 물을 받는다 — 테두리 띠에 걸친
      다른 노드(지나가는 관 등)는 세지 않는다. attach_heads_center 가 편집
      최종망에 중심 연결을 만들어 두므로, 잣대도 그 확정을 따른다.
      (가지관을 끊으면 헤드가 말라야 한다 — 오너 절단 시험 2026-08-13)
    """
    tol = HEAD_TOUCH if tol is None else float(tol)
    ng = defaultdict(list)
    for i, (x, y) in enumerate(pts):
        gput(ng, 500.0, x, y, i)
    ups = {(round(x, 1), round(y, 1)) for (x, y, _r) in upright}
    eg = None
    if edges and ups:
        eg = defaultdict(list)
        for (i, j) in edges:
            a, b = pts[i], pts[j]
            n = max(1, int(math.hypot(b[0] - a[0], b[1] - a[1]) / 500.0))
            for k in range(n + 1):
                t = k / n
                gput(eg, 500.0, a[0] + (b[0] - a[0]) * t,
                        a[1] + (b[1] - a[1]) * t, (i, j))
    deg = None
    if edges:
        deg = defaultdict(int)
        for (i, j) in edges:
            deg[i] += 1
            deg[j] += 1
    out = []
    for (hx, hy, hr) in hcov:
        lim = hr + tol
        near = set(gnear(ng, 500.0, hx, hy, rings=1 + int(lim // 500)))
        keep = set()
        for n in near:
            d = math.hypot(pts[n][0] - hx, pts[n][1] - hy)
            if d <= ARM_CTR or abs(d - hr) <= tol:
                keep.add(n)
        if deg is not None and keep:
            ctrn = []
            for n in keep:
                if deg.get(n, 0) < 1:
                    continue
                d = math.hypot(pts[n][0] - hx, pts[n][1] - hy)
                if d <= ARM_CTR:
                    ctrn.append((d, n))
            if ctrn:
                keep = {min(ctrn)[1]}
        if not keep and eg is not None \
                and (round(hx, 1), round(hy, 1)) in ups:
            rings = 1 + int(lim // 500)
            for (i, j) in set(gnear(eg, 500.0, hx, hy, rings=rings)):
                d, _t = seg_dist(pts[i], pts[j], hx, hy)
                if d <= hr:
                    keep.add(i)
                    keep.add(j)
        out.append(keep)
    return out


def attach_heads_center(pts, edges, hcov, tol=None):
    """헤드 접속 «완성» — 편집 최종망의 헤드를 중심 노드로 확정한다.

    [2026-08-13 오너] 유저편집이 끝난 망은 kfp 와 같이 «헤드 = 중심 노드,
    팔 하나로 접속»이어야 한다. 변환은 이것을 1:1 로 읽기만 한다.

    재판정 없음 — 이미 이어진 망을 조회만 한다:
      ① 중심(≤ARM_CTR)에 간선 달린 노드가 있으면 그것이 접속점.
         (상향식 5단계 원 밑 쪼갬 노드 · 중심까지 그려진 팔 포함)
      ② 없으면 테두리 띠(±tol)에서 «선이 끝나는» 노드(간선 1개) 중
         중심 최근접 끝을 헤드 중심 새 노드와 잇는다. 이 마지막 한 뼘은
         배관 창작이 아니라 헤드 자신의 목이다 [오너 확정 2026-08-13].
         단, 반대쪽 끝도 같은 원 테두리(±2mm)에 앉은 «현»은 문양(하향
         기호 가로막대 — netskip 이 토막일 수 있어 살려 둔 것)이지
         팔이 아니다 — 팔은 원 밖에서 들어온다. 3F 헤드#169 실측:
         막대 끝 d=148 이 진짜 팔 끝 d=150 보다 2mm 가까워 오집던 자리.
      ③ 둘 다 없으면 잇지 않는다 — 지나가는 배관은 상향식만 접속한다.
         (상향식 원 밑 통과 젖음은 head_nodes 의 기존 문이 그대로 담당)

    idempotent — 같은 망에 두 번 돌려도 결과가 같다. 유저가 팔(짧은 선)을
    지우면 끝이 사라지므로 다시 잇지 않는다.

    반환: (pts, edges, centers, n_wire, multi)
      centers[i] = i번째 헤드의 중심 노드 or None
      n_wire     = 이번에 새로 이은 끝→중심 수
      multi      = 끝이 박빙(당선·차선 차 < ARM_TIE)이던 헤드 인덱스 목록 (보고용)
    """
    tol = HEAD_TOUCH if tol is None else float(tol)
    pts2 = list(pts)
    edges2 = {tuple(sorted(e)) for e in edges}
    nbr = defaultdict(set)
    for i, j in edges2:
        nbr[i].add(j)
        nbr[j].add(i)
    ng = defaultdict(list)
    for i, (x, y) in enumerate(pts2):
        gput(ng, 500.0, x, y, i)
    centers, multi = [], []
    n_wire = 0
    for di, (hx, hy, hr) in enumerate(hcov):
        lim = hr + tol
        near = set(gnear(ng, 500.0, hx, hy, rings=1 + int(lim // 500)))
        ctr, ends = [], []
        for n in near:
            if not nbr.get(n):
                continue
            d = math.hypot(pts2[n][0] - hx, pts2[n][1] - hy)
            if d <= ARM_CTR:
                ctr.append((d, n))
            elif abs(d - hr) <= tol and len(nbr[n]) == 1:
                o = next(iter(nbr[n]))
                do = math.hypot(pts2[o][0] - hx, pts2[o][1] - hy)
                if abs(do - hr) <= 2.0:
                    continue    # 현(양끝이 테두리) = 문양 — 팔이 아니다
                ends.append((d, n))
        if ctr:
            centers.append(min(ctr)[1])
            continue
        if not ends:
            centers.append(None)
            continue
        if len(ends) > 1:
            ds = sorted(d for d, _n in ends)
            if ds[1] - ds[0] < ARM_TIE:
                multi.append(di)
        e = min(ends)[1]
        c = len(pts2)
        pts2.append((float(hx), float(hy)))
        gput(ng, 500.0, float(hx), float(hy), c)
        edges2.add(tuple(sorted((e, c))))
        nbr[e].add(c)
        nbr[c].add(e)
        centers.append(c)
        n_wire += 1
    return pts2, frozenset(edges2), centers, n_wire, multi


def flow(pts, edges0, spots, arms, node_spots, seed, answers=()):
    """물길 — 관 따라 가고, 표시 만나면 통과·분기, 없으면 거기서 막힌다.

    edges0 를 **건드리지 않는다** — 답을 하나 넣어 보고 되돌리며 값어치를
    재려면 물길을 몇 번이고 다시 돌려야 한다.
    answers = 유저가 고른 이음 [(노드, 노드), ...] — 표시가 없어 막힌 자리다.
    """
    edges = set(edges0)
    adj = defaultdict(set)
    for (i, j) in edges:
        adj[i].add(j)
        adj[j].add(i)
    deg0 = {i: len(adj[i]) for i in adj}
    ans = defaultdict(set)
    for (u, v) in answers:
        ans[u].add(v)
        ans[v].add(u)

    reach, q = set(seed), deque(seed)
    used, joins, asks, ends = set(), [], [], []
    passed = Counter()
    while q:
        u = q.popleft()
        for v in list(adj[u]):
            if v not in reach:
                reach.add(v)
                q.append(v)
        opened = 0
        term = False
        for si in node_spots.get(u, ()):
            # ★헤드는 언제나 관말이다 — 팔이 몇 개로 보이든 통과시키지 않는다
            #   [오너 확정 2026-08-07 밤]. 물은 여기서 멈추고, 묻지도 않는다.
            if spots[si]["k"] == "헤드":
                term = True
                continue
            others = [n for n in arms[si] if n != u]
            if not others:
                continue
            opened += 1
            if si not in used:
                used.add(si)
                passed[spots[si]["k"]] += 1
            for n in others:
                key = (min(u, n), max(u, n))
                if key not in edges:
                    edges.add(key)
                    joins.append((spots[si]["k"], u, n))
                adj[u].add(n)
                adj[n].add(u)
                if n not in reach:
                    reach.add(n)
                    q.append(n)
        for n in ans.get(u, ()):
            key = (min(u, n), max(u, n))
            if key not in edges:
                edges.add(key)
                joins.append(("오너답", u, n))
            adj[u].add(n)
            adj[n].add(u)
            if n not in reach:
                reach.add(n)
                q.append(n)
        if not opened and not ans.get(u) and deg0.get(u, 0) <= 1:
            if term:
                ends.append(u)
            else:
                asks.append(u)
    return dict(reach=reach, edges=edges, joins=joins, asks=asks,
                ends=ends, passed=passed)


def _join_order(pts, nodes, cx, cy):
    """접속부에서 팔 잇는 순서 — .cursorrules «접속부 잇는 정석» 그대로.

      ① 가로 배관(표시 중심이 선분 «사이»에 있는 팔 쌍)을 먼저 잇는다
      ② 허브(중심에 가장 가까운 팔)를 가로에 묶는다
      ③ 남는 팔(헤드로 가는 짧은 선 등)은 허브에 붙인다

    예전: `base = a[0]` 에 나머지를 전부 붙이는 별 모양.
    첫 팔이 짧은 선 끝이면 좌우 가로배관이 그 점으로 45도 끌려 올라갔다.

    ★통과 가로는 중심이 선분 «사이»에 있을 때만 인정한다 [2026-08-07 밤].
      짧은선↔중심노드 처럼 중심이 끝점에만 있는 쌍을 두 번째 가로로 잡으면
      가로는 이어지고 세로·중심만 동떨어져 빨간 점이 된다(3F 실측).
      여유는 기존 `ARM_CTR` 를 그대로 쓴다. 짧은 선이 아주 짧아도 되고,
      세로선이 아예 없으면 가로만 잇고 헤드는 안 붙은 채로 둔다(선을 안 만듦).
    """
    nodes = list(dict.fromkeys(nodes))
    if len(nodes) < 2:
        return []
    if len(nodes) == 2:
        return [(nodes[0], nodes[1])]

    def chord(u, v):
        """중심이 선분 uv «사이»에 있으면 (중심까지거리, -길이, u, v)."""
        ax, ay = pts[u][0], pts[u][1]
        bx, by = pts[v][0], pts[v][1]
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        if L2 < 1e-18:
            return None
        L = math.sqrt(L2)
        t = ((cx - ax) * dx + (cy - ay) * dy) / L2
        # 끝점에만 중심이 있으면 짧은선↔중심. 통과 가로가 아니다.
        edge = min(0.45, max(ARM_CTR / L, 1e-6))
        if t < edge or t > 1.0 - edge:
            return None
        d = math.hypot(ax + t * dx - cx, ay + t * dy - cy)
        return (d, -L, u, v)

    # ① 통과 가로 — 티면 한 쌍, 십자면 두 쌍
    remaining = set(nodes)
    pairs = []
    while len(remaining) >= 2:
        best = None
        rl = list(remaining)
        for i, u in enumerate(rl):
            for v in rl[i + 1:]:
                sc = chord(u, v)
                if sc is None:
                    continue
                if best is None or sc < best:
                    best = sc
        if best is None or best[0] > ARM_CTR:
            break
        _d, _nL, u, v = best
        pairs.append((u, v))
        remaining.discard(u)
        remaining.discard(v)

    # ② 허브 = 모든 팔 중 중심에 가장 가까운 것 (보통 thru_arms 중심 노드)
    hub = min(nodes, key=lambda n: math.hypot(pts[n][0] - cx, pts[n][1] - cy))
    out = list(pairs)
    used = {n for uv in pairs for n in uv}
    if pairs and hub not in used:
        anchor = min(used, key=lambda n: math.hypot(
            pts[n][0] - pts[hub][0], pts[n][1] - pts[hub][1]))
        out.append((hub, anchor))
        used.add(hub)
    elif not pairs:
        used.add(hub)

    # ③ 짧은 선 등 남은 팔 → 허브에 붙인다
    for n in nodes:
        if n in used:
            continue
        out.append((hub, n))
        used.add(n)
    return out


def join_all(pts, edges0, spots, arms):
    """★물을 도면 전체에 부었다고 보고 «모든 접속표시»를 통과시킨다.

    급수원 한 점에서 흘리는 것과 달리, 표시가 있는 자리는 물이 오는지 마는지
    따지지 않고 다 통과시킨다. 끊긴 곳은 그대로 끊긴다 — 표시가 없으면 이을
    근거가 없다. 그 결과 도면이 «물덩이» 몇 개로 갈리는지가 그대로 드러난다.

    ★★헤드는 여기서 아무것도 잇지 않는다 [오너 확정 2026-08-07 밤].
      "헤드는 접속표시는 아니다. 다만 평면도에서는 그렇게 보일 수 있다."
      헤드는 관말이다. 상향식이라 배관 위에 앉은 경우조차 실제로는 티에서
      z축으로 갈라져 올라간 것이라, 평면도에 안 보이는 그 짧은 관 하나가
      헤드의 유일한 팔이다. 헤드에 팔이 둘로 보이면 그것은 도면을 잘못 읽은
      것이지 갈림길이 아니다.
      헤드를 통과점으로 쓰면 «배관이 헤드를 경유해» 이어진다 — 오너 판정으로
      수리계산 대형사고다(.cursorrules 원칙 5).

    ★잇는 순서 [오너 2026-08-07 · .cursorrules 정석]:
      가로 배관끼리 먼저 → 짧은 선은 그다음. `_join_order` 참고.
    """
    edges = set(edges0)
    joins, passed = [], Counter()
    for si, a in enumerate(arms):
        if spots[si]["k"] == "헤드":
            continue
        if len(a) < 2:
            continue
        passed[spots[si]["k"]] += 1
        sp = spots[si]
        for u, v in _join_order(pts, a, sp["cx"], sp["cy"]):
            k = (min(u, v), max(u, v))
            if k not in edges:
                edges.add(k)
                joins.append((sp["k"], u, v))
    return edges, joins, passed


# --------------------------------------------- 4단계 «끊긴 배관 잇기»
#   (옛 이름: 3단계 끊어 그린 자리 · 성적 키 옛 이음3 → 새 이음4)
def stage4_body(st, spots):
    """4단계 끊긴 배관 잇기 — `s45.join_by_through_main` [오너 2026-08-07].

    기호 없이 글자·다른 배관에 가려 **일부러 끊어 그린** 틈을 잇는다. 접속표시가
    있는 틈은 함수가 스스로 뺀다(`gap_has_symbol`) — 그래서 2단계(접속부
    이음)가 이을 자리와 겹치지 않는다.

    ★넘겨받은 그래프에 간선을 **더하므로**, 1단계 그래프가 오염되지 않게
      복사본을 넘기고 이음 목록만 받아 온다.
    ★찍은 것 밖은 안 본다(원칙 1) — 비재료 설명자 배관(`explain_segs`)을
      넘기지 않는다.
    ★노드 번호는 1단계 그래프 것이라 시제품 `pts` 와 그대로 맞는다.
    반환: ([(u, v), ...], 부수정보)
    """
    g2 = copy.deepcopy(st["g"])
    eb2 = dict(st["ebundle"])
    syms = [(s["cx"], s["cy"], s["r"], s["k"]) for s in spots]
    mat_layers = {ly for ly, _c in st["mat_bundles"]}
    joins, side = s45.join_by_through_main(
        g2, eb2, syms, st["knobs"], texts=st["w"].texts,
        mat_layers=mat_layers, explain_segs=())
    out = []
    for j in joins:
        uv = j.get("_nodes")
        if uv:
            out.append((int(uv[0]), int(uv[1])))
    return out, side


# ------------------------------------------- 5단계 «상향식 헤드 접속»
#   (옛 이름: 4단계 헤드걸침 · 성적 키 옛 이음4 → 새 이음5)
def head_kind(layer, given=None):
    """헤드 종류 — 판정/라벨이 먼저, 레이어 이름 짐작은 폴백.

    [2026-08-08] given = classify 결과 또는 옛 kind. 하향↔하향식 정규화.
    없으면 레이어 이름(구스펙·미찍힘). 그것도 없으면 미지정.
    """
    if given is not None and str(given).strip() != "":
        return normalize_head_kind(given)
    s = str(layer or "")
    if "상하향" in s:
        return "상하향식"
    if "상향" in s:
        return "상향식"
    if "하향" in s:
        return "하향식"
    return "미지정"


def _spec_pick_for(spec, bundle, r=None, tri_side=None, mark_fp=None):
    """헤드 실체에 맞는 찍기 픽(mark_fp·구 mark_bundle·label 호환).

    mark_bundles 가 있는 신규 픽은 지문 완전일치를 건너뛴다
    (문양 포함 매칭은 classify_head_kind 가 담당) [2026-08-09].
    """
    b = tuple(bundle)
    best = None
    want_fp = (None if mark_fp is None
               else heads.mark_fp_key(mark_fp))
    for hs in (spec.get("heads") or ()):
        if tuple(hs["bundle"]) != b:
            continue
        if tri_side is not None and "tri_side" in hs:
            side = float(hs["tri_side"])
            if abs(side - float(tri_side)) <= max(5.0, side * 0.10):
                return hs
        if r is not None and "r" in hs:
            rr = float(hs["r"])
            gap = abs(rr - float(r))
            if gap > max(5.0, rr * 0.10):
                continue
            if want_fp is not None and hs.get("mark_fp") is not None:
                # 신규 문양묶음 픽 — 완전일치 대신 묶음×r 최선
                if hs.get("mark_bundles") is not None:
                    if best is None or gap < best[0]:
                        best = (gap, hs)
                    continue
                if heads.mark_fp_key(hs["mark_fp"]) != want_fp:
                    continue
                return hs
            if best is None or gap < best[0]:
                best = (gap, hs)
    return None if best is None else best[1]


def dual_marks_of(spec):
    """구호환 — dual_marks + heads[].mark_bundle. 신규 스펙은 비는 것이 정상."""
    out, seen = [], set()
    for mb in (spec.get("dual_marks") or ()):
        t = tuple(mb)
        if t in seen:
            continue
        seen.add(t)
        out.append(t)
    for hs in (spec.get("heads") or ()):
        if hs.get("mark_bundle") is None:
            continue
        t = tuple(hs["mark_bundle"])
        if t in seen:
            continue
        seen.add(t)
        out.append(t)
    return out


def _owned_half_arc_head_keys(pts, edges, spots, arms):
    """팔의 종단 선 끝이 2단계 반호 중심에 닿는 헤드 → {disk_key: 팔 단위방향}.

    같은 좌표의 CAD 분할 노드는 합치고, 헤드 내부선은 시작에서 제외한다.
    방향 전환과 길이는 제한하지 않되 분기나 다른 접속표시를 넘어가지 않는다.
    반호 원주 arm은 소유 근거가 아니며, 종단점이 실제 ARC 중심인 경우만 받는다.
    값은 반호로 가는 첫 진출 방향(통과 배관 판정용).
    """
    parent = list(range(len(pts)))

    def find(n):
        while parent[n] != n:
            parent[n] = parent[parent[n]]
            n = parent[n]
        return n

    def union(a, b):
        a, b = find(a), find(b)
        if a != b:
            parent[b] = a

    for u, v in edges:
        if math.hypot(pts[u][0] - pts[v][0], pts[u][1] - pts[v][1]) <= 1e-9:
            union(u, v)

    adj = defaultdict(set)
    for u, v in edges:
        u, v = find(u), find(v)
        if u == v:
            continue
        adj[u].add(v)
        adj[v].add(u)

    arc_center_nodes = set()
    for si, ns in enumerate(arms):
        spot = spots[si]
        if spot["k"] != "호":
            continue
        for n in ns:
            if math.hypot(
                    pts[n][0] - spot["cx"],
                    pts[n][1] - spot["cy"]) <= s1.SNAP:
                arc_center_nodes.add(find(n))

    found = {}
    for si, ns in enumerate(arms):
        spot = spots[si]
        if spot["k"] != "헤드":
            continue
        hx, hy, hr = spot["cx"], spot["cy"], spot["r"]
        starts = {find(n) for n in ns}
        todo = []
        for u in starts:
            for v in adj.get(u, ()):
                # 양 끝이 헤드 원 안이면 무늬선/잡선이다.
                if (math.hypot(pts[u][0] - hx, pts[u][1] - hy) <= hr
                        and math.hypot(pts[v][0] - hx, pts[v][1] - hy) <= hr):
                    continue
                todo.append((u, v, u, v))  # prev, node, first_u, first_v
        seen = set()
        while todo:
            prev, node, fu, fv = todo.pop()
            state = (prev, node)
            if state in seen:
                continue
            seen.add(state)
            # 반호 중심에서 끝난 선만 이 헤드가 소유한 접속기호다.
            if node in arc_center_nodes:
                d0 = math.hypot(pts[fu][0] - hx, pts[fu][1] - hy)
                d1 = math.hypot(pts[fv][0] - hx, pts[fv][1] - hy)
                ox, oy = ((pts[fv][0], pts[fv][1]) if d1 >= d0
                          else (pts[fu][0], pts[fu][1]))
                d = math.hypot(ox - hx, oy - hy) or 1.0
                found[_disk_key(hx, hy, hr)] = ((ox - hx) / d, (oy - hy) / d)
                break
            # 분기/종단을 넘어가지 않는다. 접속표시 틈은 edges 자체가 끊겨 있다.
            nxt = [n for n in adj.get(node, ()) if n != prev]
            if len(nxt) != 1:
                continue
            todo.append((node, nxt[0], fu, fv))
    return found


def _head_has_opposite_material(w, mat_set, hx, hy, hr, index=None,
                                along_dir=None):
    """헤드 테두리 맞은편 재료 — along_dir 있으면 그 팔축 일직선 통과만.

    양쪽 터치점이 헤드 중심·팔축과 거의 일직선(횡이탈 ≤ ARM_CTR)일 때만
    True. Y만 어긋난 좌우 거리터치는 통과로 보지 않는다.
    """
    ctr = heads._ARM_CTR
    touch = heads._HEAD_TOUCH
    mat_set = set(tuple(b) for b in mat_set)
    hits = []  # (outer_ux, outer_uy, px, py)
    if index is None:
        rows = ((order, a, b, math.hypot(b[0] - a[0], b[1] - a[1]))
                for order, (ly, c, a, b) in enumerate(w.segs)
                if (ly, c) in mat_set)
    else:
        rows = heads._head_arm_candidates(
            index, hx, hy, max(ctr, float(hr) + touch))
    for _order, a, b, ln in rows:
        if ln <= 0:
            continue
        for (px, py), (ox, oy) in ((a, b), (b, a)):
            d = math.hypot(px - hx, py - hy)
            if not (d <= ctr or abs(d - hr) <= touch):
                continue
            od = math.hypot(ox - hx, oy - hy)
            if od <= hr + touch:
                continue
            hits.append(((ox - hx) / od, (oy - hy) / od, px, py))
    ax = ay = None
    if along_dir is not None:
        ax, ay = float(along_dir[0]), float(along_dir[1])
    for i, (ux, uy, px, py) in enumerate(hits):
        for vx, vy, qx, qy in hits[i + 1:]:
            if ux * vx + uy * vy >= -0.5:
                continue
            if ax is None:
                # along 없으면 중심→터치점 방향이 거의 반대일 때만
                d1 = math.hypot(px - hx, py - hy) or 1.0
                d2 = math.hypot(qx - hx, qy - hy) or 1.0
                nux, nuy = (px - hx) / d1, (py - hy) / d1
                nvx, nvy = (qx - hx) / d2, (qy - hy) / d2
                if nux * nvx + nuy * nvy <= -0.85:
                    return True
                continue
            # 반호 팔 축 정렬 + 양쪽 터치점이 그 축 일직선 위
            if ux * ax + uy * ay < 0.85 and vx * ax + vy * ay < 0.85:
                continue
            lat_p = abs((px - hx) * ay - (py - hy) * ax)
            lat_q = abs((qx - hx) * ay - (qy - hy) * ax)
            if lat_p <= ctr and lat_q <= ctr:
                return True
    return False


def _kind_graph(st):
    """종류 판정용 그래프. 작업 망에 안 넣은 문양 막대기만 복사본에 붙인다."""
    bars = st.get("head_symbol_bars") or ()
    if not bars:
        return st["g"]
    g2 = copy.deepcopy(st["g"])
    s1.attach_head_symbol_bars(g2, bars)
    return g2


def _owned_half_arc_keys_from_stage1(st):
    """단독 분류 호출용. pipeline은 이미 만든 1·2단계 자료를 넘긴다."""
    cache_key = "_owned_half_arc_head_keys"
    if cache_key not in st:
        g = _kind_graph(st)
        spots, _hcov = spots_body(st, owner=True)
        arms, _ns = spot_arms(g, spots)
        pts, edges, arms, _n_thru = thru_arms(
            g.pts, frozenset(g.edges), spots, arms)
        st[cache_key] = _owned_half_arc_head_keys(pts, edges, spots, arms)
    return st[cache_key]


def classify_head_kind(st, head, fp_spatial=None, arm_index=None,
                       owned_half_arc=None):
    """헤드마다 종류 [2026-08-08 오너 확정 · 문양 지문].

      1) 헤드 원 fp 가 x칸(상하향) 픽과 같으면 → 상하향식
      2) 구호환: dual_marks/mark_bundle 도형이 원 안 → 상하향식
      3) 구스펙: 칸 상하향 + mark 없음 → 픽 전체 상하향식
      4) tri_side → 하향식
      5) 팔 종단이 소유 반호 중심 → 하향식
         (단, 그 팔 축으로 맞은편·통과 배관 있으면 상향식)
      6) 소유 반호 없음 → 상향식

    owned_half_arc: None=조회 / False=없음 / True=있음 /
    (ux,uy)=있음+반호 팔 방향.
    """
    w, kn = st["w"], st["knobs"]
    mat_set = set(st["mat_bundles"])
    small_len = kn["small_len"]
    spec = st["spec"]
    b = tuple(head.get("bundle") or ())
    if "tri_side" in head:
        return "하향식"
    if "head_r" not in head:
        return "미지정"
    hx, hy = head["c"][0], head["c"][1]
    hr = float(head["head_r"])
    fp = heads.mark_fingerprint(
        w, hx, hy, hr, small_len, spatial=fp_spatial)
    fp_key = heads.mark_fp_key(fp)
    mat_set = set(tuple(t) for t in (spec.get("material_picks") or ()))
    # 1) x칸 픽과 문양 일치 → 상하향식
    #    mark_bundles 있으면 재료확정·포함매칭 [2026-08-09],
    #    없으면 옛 완전일치.
    for hs in (spec.get("heads") or ()):
        if tuple(hs.get("bundle") or ()) != b or "r" not in hs:
            continue
        if normalize_head_slot(hs.get("label")) != "상하향":
            continue
        rr = float(hs["r"])
        if abs(rr - hr) > max(5.0, rr * 0.10):
            continue
        hfp = hs.get("mark_fp")
        if hfp is None:
            continue
        mb = hs.get("mark_bundles")
        if mb is not None and mat_set:
            got = heads.mark_fp_on_bundles(
                w, hx, hy, hr, small_len, mb, spatial=fp_spatial)
            if heads.mark_fp_contains(hfp, got):
                return "상하향식"
            continue
        if heads.mark_fp_key(hfp) == fp_key:
            return "상하향식"
    # 2) 구호환 dual_marks / mark_bundle
    marks = dual_marks_of(spec)
    if marks:
        for mb in marks:
            if heads.mark_bundle_in_disk(w, mb, hx, hy, hr, small_len):
                return "상하향식"
    pick = _spec_pick_for(spec, b, r=hr, mark_fp=fp)
    # 3) 구스펙: label 상하향 + mark_fp/mark_bundle 없음 → 픽 전체
    pick_mark = None
    if pick and pick.get("mark_bundle") is not None:
        pick_mark = tuple(pick["mark_bundle"])
    elif head.get("mark_bundle") is not None:
        pick_mark = tuple(head["mark_bundle"])
    slot_src = (pick or {}).get("label", head.get("label"))
    if (normalize_head_slot(slot_src) == "상하향"
            and pick_mark is None
            and (pick or {}).get("mark_fp") is None):
        return "상하향식"
    # 5~6) 소유 반호 → 하향. 반호 팔 축 통과 배관만 상향으로 뒤집는다.
    along_dir = None
    hkey = _disk_key(hx, hy, hr)
    if owned_half_arc is None:
        along_dir = _owned_half_arc_keys_from_stage1(st).get(hkey)
        owned = along_dir is not None
    elif owned_half_arc is False:
        owned = False
    elif owned_half_arc is True:
        along_dir = _owned_half_arc_keys_from_stage1(st).get(hkey)
        owned = True
    else:
        along_dir = owned_half_arc
        owned = True
    if owned and along_dir is not None and _head_has_opposite_material(
            w, mat_set, hx, hy, hr, index=arm_index, along_dir=along_dir):
        return "상향식"
    return "하향식" if owned else "상향식"


def kind_of_head(st, head):
    """classify + 레이어 이름 폴백 — upright/물길용."""
    k = classify_head_kind(st, head)
    if k != "미지정":
        return k
    return head_kind((head.get("bundle") or ("", None))[0], None)


def stage11_classify_heads(st, arm_index=None, owned_half_arc_keys=None):
    """1-1 헤드 종류 분류 — 기존 classify_head_kind 재사용.

    추출·저장·미지정 알림. 5단계는 반환된 head_kinds 를 후보에 쓴다
    (미지정은 알리기만 — 5단계에서 새로 빼지 않음).
    반환: 헤드마다 c/bundle/kind (+ head_r 또는 tri_side).
    """
    w, spec, knobs = st["w"], st["spec"], st["knobs"]
    clusters0 = heads.collect_head_clusters(w, spec, knobs)
    head_cls, _marks, _info = heads.split_head_circles(clusters0, knobs)
    fp_spatial = heads._fp_build_spatial(w, knobs["small_len"])
    if arm_index is None:
        arm_index = heads.build_head_arm_index(
            w, spec.get("material_picks") or ())
    if owned_half_arc_keys is None:
        owned_half_arc_keys = _owned_half_arc_keys_from_stage1(st)
    out = []
    for h in head_cls:
        hkey = (_disk_key(h["c"][0], h["c"][1], h["head_r"])
                if "head_r" in h else None)
        if hkey is None:
            owned_arg = None
        elif isinstance(owned_half_arc_keys, dict):
            owned_arg = owned_half_arc_keys.get(hkey, False)
        else:
            owned_arg = hkey in owned_half_arc_keys
        kind = classify_head_kind(
            st, h, fp_spatial=fp_spatial, arm_index=arm_index,
            owned_half_arc=owned_arg)
        rec = {
            "c": (float(h["c"][0]), float(h["c"][1])),
            "bundle": tuple(h.get("bundle") or ()),
            "kind": kind,
        }
        if "head_r" in h:
            rec["head_r"] = float(h["head_r"])
        if "tri_side" in h:
            rec["tri_side"] = float(h["tri_side"])
        out.append(rec)
    counts = Counter(r["kind"] for r in out)
    label = STAGE_NAME.get("1-1", "헤드 종류 분류")
    parts = " · ".join(f"{k} {v}" for k, v in counts.most_common())
    print(f"\n[1-1 {label}] 헤드 {len(out)}개"
          + (f" · {parts}" if parts else ""))
    unknown = [r for r in out if r["kind"] == "미지정"]
    if unknown:
        print(f"    ★미지정 {len(unknown)}개 — 알리기만 (이후 단계 제외 안 함)")
        for r in unknown[:40]:
            x, y = r["c"]
            extra = ""
            if "head_r" in r:
                extra = f" r={r['head_r']:.1f}"
            elif "tri_side" in r:
                extra = f" tri={r['tri_side']:.1f}"
            b = r.get("bundle") or ()
            layer = b[0] if b else "?"
            print(f"      ({x:.1f}, {y:.1f}){extra} · {layer}")
        if len(unknown) > 40:
            print(f"      … 외 {len(unknown) - 40}개")
    else:
        print("    미지정 0개")
    return out


def _dual_ok_for_stage5(st, hx, hy, hr, arm_index=None):
    """상하향식 5단계 후보 — 팔이면 제외, 그 외(관위·기타)는 포함.

    팔이 있으면 하향처럼 재료 팔로 이미 붙으므로 join_by_head_cover 에
    넣지 않는다. 새 이음 로직은 만들지 않는다 [2026-08-09 오너].
    종류 재분류는 하지 않고 부착 기하만 본다.
    """
    w, kn = st["w"], st["knobs"]
    mat_set = set(tuple(t) for t in (st["spec"].get("material_picks") or ()))
    small_len = kn["small_len"]
    if heads.head_has_arm(
            w, mat_set, hx, hy, hr, small_len, index=arm_index):
        return False
    return True


def upright_disks(st, hcov, head_kinds, arm_index=None):
    """5단계 join_by_head_cover 후보 [오너 2026-08-09 · 1-1 kind 소비].

    종류는 1-1 head_kinds 를 쓰고 다시 classify 하지 않는다.
      상향식 INCLUDE · 하향식 EXCLUDE
      상하향식 + 팔 EXCLUDE · 상하향식 + 관위(비팔) INCLUDE
      미지정 INCLUDE (1-1에서 알림만 — 여기서 새로 빼지 않음)
    «관이 헤드를 관통하다 원 자리만 비었다»는 join_by_head_cover 가 가른다.
    """
    by = {}
    if arm_index is None:
        arm_index = heads.build_head_arm_index(
            st["w"], st["spec"].get("material_picks") or ())
    for rec in head_kinds or ():
        if "head_r" not in rec:
            continue
        by[_disk_key(rec["c"][0], rec["c"][1], rec["head_r"])] = rec
    out = []
    n_dual_arm = 0
    n_miss = 0
    for (hx, hy, hr) in hcov:
        rec = by.get(_disk_key(hx, hy, hr))
        if rec is None:
            # 확정 kind 없는 디스크를 조용히 후보에 넣지 않음
            n_miss += 1
            continue
        kind = normalize_head_kind(rec.get("kind"))
        if kind == "하향식":
            continue
        if kind == "상하향식":
            if not _dual_ok_for_stage5(
                    st, hx, hy, hr, arm_index=arm_index):
                n_dual_arm += 1
                continue
        # 상향식 · 미지정 · 상하향식(비팔)
        out.append((hx, hy, hr))
    if n_dual_arm or n_miss:
        print(f"    [5 후보] 상하향팔제외 {n_dual_arm}"
              + (f" · hcov미매칭제외 {n_miss}" if n_miss else ""))
    return out


def stage5_body(st, ups):
    """5단계 상향식 헤드 접속 — ① 양쪽 틈 이음(`join_by_head_cover`).

    조건 둘을 다 만족해야 잇는다 — ① 헤드 원이 틈 축을 자를 것 ② 원 테두리
    밖으로 남는 길이가 양 끝 각각 150mm 이하일 것. 즉 «관이 원을 관통하다
    원 자리만 비워졌다»는 모양이다. 헤드를 통과점으로 쓰지 않고 **메인끼리**만.

    ② 원 밑 통과관 중심 분할은 `stage5_split_through_uprights` (같은 5단계).
    ★2·4단계와 «섞지 않고» 따로 센다 [오너 이식 규칙]. 성적표에 「이음5」.
    반환: [(u, v), ...]
    """
    if not ups:
        return []
    g2 = copy.deepcopy(st["g"])
    eb2 = dict(st["ebundle"])
    joins = s45.join_by_head_cover(g2, eb2, ups, st["knobs"])
    out = []
    for j in joins:
        uv = j.get("_nodes")
        if uv:
            out.append((int(uv[0]), int(uv[1])))
    return out


def stage5_split_through_uprights(pts, edges, ups):
    """5단계 ② — 상향식 · 원 밑으로만 지나가는 관을 중심에서 쪼갠다.

    [2026-08-12 오너] MF101처럼 원 양쪽에 관 끝이 있으면 ① 틈이음으로
    중심 접속이 된다. 원 밑 통과만 있고 중심/테두리 노드가 없으면 KFP 변환
    때 먼 끝 노드에 헤드가 겹쳐 사라진다. 같은 5단계에서 중심 노드를 만들어
    통일한다. 높이(Z)는 주지 않아도 평면 접속은 된다.

    · 대상: ups 헤드 중, 그래프에 중심(≤ARM_CTR)·테두리(±HEAD_TOUCH) 노드가
      없고, 원 안(횡이탈 ≤ r)을 관이 지나는 것
    · 동작: 통과 선분을 헤드 중심에서 분할 → 새 노드 (좌·우 관이 경유)
    · 하향식·이미 테두리 붙은 상향식은 손대지 않음
    반환: (pts2, edges2, n_split)
    """
    if not ups:
        return list(pts), set(tuple(sorted(e)) for e in edges), 0
    pts2 = list(pts)
    edges2 = {tuple(sorted(e)) for e in edges}
    used = {n for e in edges2 for n in e}

    def has_near_node(hx, hy, hr):
        for n in used:
            d = math.hypot(pts2[n][0] - hx, pts2[n][1] - hy)
            if d <= ARM_CTR or abs(d - hr) <= HEAD_TOUCH:
                return True
        return False

    # 통과 후보: 헤드 → (횡이탈, |t-0.5|, i, j, t)
    want = []
    cell = 1000.0
    eg = defaultdict(list)
    for (i, j) in edges2:
        a, b = pts2[i], pts2[j]
        n = max(1, int(math.hypot(b[0] - a[0], b[1] - a[1]) / cell))
        for k in range(n + 1):
            t = k / n
            gput(eg, cell, a[0] + (b[0] - a[0]) * t,
                    a[1] + (b[1] - a[1]) * t, (i, j))
    for (hx, hy, hr) in ups:
        if hr <= 0 or has_near_node(hx, hy, hr):
            continue
        best = None
        rings = 1 + int(hr // cell)
        for (i, j) in set(gnear(eg, cell, hx, hy, rings=rings)):
            if (min(i, j), max(i, j)) not in edges2:
                continue
            ax, ay = pts2[i]
            bx, by = pts2[j]
            L = math.hypot(bx - ax, by - ay)
            if L < 1e-9:
                continue
            t = ((hx - ax) * (bx - ax) + (hy - ay) * (by - ay)) / (L * L)
            # 끝점에 앉은 것은 이미 테두리/중심 후보 — 여긴 통과만
            if t <= 1e-6 or t >= 1.0 - 1e-6:
                continue
            px, py = ax + (bx - ax) * t, ay + (by - ay) * t
            lat = math.hypot(px - hx, py - hy)
            if lat > hr:
                continue
            sc = (lat, abs(t - 0.5))
            if best is None or sc < best[0]:
                best = (sc, i, j, t)
        if best is not None:
            want.append((hx, hy, best[1], best[2], best[3]))

    # 같은 선분에 헤드가 여럿이면 t 순으로 한 번에 쪼갠다 (thru_arms 와 동일)
    by_edge = defaultdict(list)
    for hx, hy, i, j, t in want:
        a, b = min(i, j), max(i, j)
        # t 는 (i→j) 기준. a=min 으로 정규화
        if i > j:
            t = 1.0 - t
        by_edge[(a, b)].append((t, hx, hy))

    n_split = 0
    for (i, j), lst in by_edge.items():
        if (i, j) not in edges2:
            continue
        edges2.discard((i, j))
        ax, ay = pts2[i]
        bx, by = pts2[j]
        L = math.hypot(bx - ax, by - ay)
        lst.sort()
        prev = i
        k = 0
        while k < len(lst):
            t0 = lst[k][0]
            # 같은 자리(1mm 안)는 노드 하나 — 헤드 중심은 첫 좌표
            hx, hy = lst[k][1], lst[k][2]
            while k < len(lst) and abs(lst[k][0] - t0) * L <= 1.0:
                k += 1
            # 중심 좌표에 노드 (lat≈0 실측 · 오너: 헤드 중심 접속)
            pts2.append((float(hx), float(hy)))
            new = len(pts2) - 1
            edges2.add((min(prev, new), max(prev, new)))
            used.add(new)
            prev = new
            n_split += 1
        edges2.add((min(prev, j), max(prev, j)))
    return pts2, edges2, n_split


def add_no_ring(edges, pairs):
    """이음을 얹되 «이미 한 덩어리»인 곳은 버린다 — 도면에 없는 고리 금지.

    본체 함수도 같은 규칙을 갖고 있지만 1단계 덩어리만 보고 판정한다. 시제품은
    2단계에서 이미 많이 붙여 놓았으므로 그 상태로 다시 걸러야 한다.
    반환: (새 edges, 얹은 이음, 고리라 버린 이음)
    """
    par = {}

    def find(x):
        par.setdefault(x, x)
        while par[x] != x:
            par[x] = par[par[x]]
            x = par[x]
        return x

    e2 = set(edges)
    for i, j in e2:
        ri, rj = find(i), find(j)
        if ri != rj:
            par[ri] = rj
    took, ring = [], []
    for (u, v) in pairs:
        k = (min(u, v), max(u, v))
        if k in e2:
            continue
        ru, rv = find(u), find(v)
        if ru == rv:
            ring.append((u, v))
            continue
        par[ru] = rv
        e2.add(k)
        took.append((u, v))
    return e2, took, ring


def pipeline(st, outside=False, stage4=True, stage5=True, key=None):
    """1→1-1→2→3→4→5 (이음) · 6(물길)은 호출측.

    성적표·그림·물길이 전부 이 함수를 쓴다. 단계마다 이음을 **따로 세어**
    어느 단계가 무엇을 했는지 성적표에서 바로 갈리게 한다.

    ★번호 주의 [2026-08-08]: 옛 j3(끊어그린)→새 j4, 옛 j4(상향식)→새 j5.
      새 j3 = 헤드 접속(경유 이음 목록) — 항상 빈 목록 · head_joins=0.

    key 가 있고 `{key}_유저손질.json` 이 있으면 5단계 뒤 joins→deletes 적용.
    key is None 이면 손질 미적용 (`build_board` 이중 적용 방지).
    j2~j5 개수는 자동 이음 성적이라 손질로 바꾸지 않는다.
    """
    g = st["g"]
    if st["spec"].get("dual_marks"):
        print(f"  [찍기] 옛 dual_marks {len(st['spec']['dual_marks'])}개"
              " — 신규 저장 안 함 · 분류 폴백만 사용")
    spots, hcov = spots_body(st, owner=True, outside=outside)
    g_kind = _kind_graph(st)
    arms_k, _ns_k = spot_arms(g_kind, spots)
    _pts_k, e1_k, arms_k, _n_k = thru_arms(
        g_kind.pts, frozenset(g_kind.edges), spots, arms_k)
    owned_half_arc_keys = _owned_half_arc_head_keys(
        _pts_k, e1_k, spots, arms_k)
    st["_owned_half_arc_head_keys"] = owned_half_arc_keys
    if g_kind is g:
        arms, _ns = arms_k, _ns_k
        pts, e1, n_thru = _pts_k, e1_k, _n_k
    else:
        arms, _ns = spot_arms(g, spots)
        pts, e1, arms, n_thru = thru_arms(
            g.pts, frozenset(g.edges), spots, arms)
    node_spots = defaultdict(list)
    for si, a in enumerate(arms):
        for n in a:
            node_spots[n].append(si)
    # 1-1 헤드 종류 분류 (추출·알림 · 5단계 후보는 이 kind 를 소비)
    arm_index = heads.build_head_arm_index(
        st["w"], st["spec"].get("material_picks") or ())
    head_kinds = stage11_classify_heads(
        st, arm_index=arm_index, owned_half_arc_keys=owned_half_arc_keys)
    # 2 접속부에서 배관 연결 (헤드 원 통과 이음 안 함 — join_all 이 헤드 skip)
    edges, j2, passed = join_all(pts, e1, spots, arms)
    # 3 헤드 접속 — 경유 이음 금지. 목록은 항상 비움(안전핀 head_joins=0).
    j3 = []
    # 4 끊긴 배관 잇기 (옛 이음3 / 옛 stage3_body)
    j4, ring4, side4 = [], [], None
    if stage4:
        pairs, side4 = stage4_body(st, spots)
        edges, j4, ring4 = add_no_ring(edges, pairs)
    # 5 상향식 헤드 접속 (옛 이음4 / 옛 stage4_body) — 후보만 head_kinds 로
    #   ① 원 양쪽 틈 이음  ② 원 밑 통과관 중심 분할 [2026-08-12 오너]
    ups = upright_disks(
        st, hcov, head_kinds, arm_index=arm_index) if stage5 else []
    j5, ring5 = [], []
    n_up_split = 0
    if stage5:
        edges, j5, ring5 = add_no_ring(edges, stage5_body(st, ups))
        pts, edges, n_up_split = stage5_split_through_uprights(pts, edges, ups)
        if n_up_split:
            print(f"    [5 통과관분할] 중심 노드 {n_up_split}개"
                  f" (상향식 원 밑 통과 → 헤드 중심 접속)")
    # 6 입구 손질 — key 를 아는 호출측만. UI(build_board)는 key=None.
    # kind_overrides 는 이음5/ups 뒤 · 색·집계용 head_kinds 만 덮는다(이음 재계산 없음).
    user_sources = []
    if key is not None:
        from services.cad_import.kinds import resolve_head_kinds
        from services.cad_import.pipeline.user_net import apply_user_edits
        pts, edges, user_sources, kind_ovs = apply_user_edits(
            key, pts, edges, default_edits_dir())
        if kind_ovs:
            # require→apply — 뒤집으면 레코드 없던 헤드에 찍은 결정이 사라진다.
            head_kinds = resolve_head_kinds(hcov, head_kinds, kind_ovs)
            counts = Counter(normalize_head_kind(r.get("kind"))
                             for r in head_kinds)
            parts = " · ".join(f"{k} {v}" for k, v in counts.most_common())
            print(f"[6 입구 손질] 헤드 종류 override {len(kind_ovs)}개 적용"
                  f" · 헤드 {len(head_kinds)}개"
                  + (f" · {parts}" if parts else ""))
    head_kinds = require_head_kinds(hcov, head_kinds)
    # 젖음 잣대의 upright 집합 = 5단계 후보와 동일
    hnodes = head_nodes(pts, hcov, edges=edges, upright=ups)
    hspot = {n for si, a in enumerate(arms) if spots[si]["k"] == "헤드"
             for n in a}
    # ★안전핀 — 「헤드를 경유해 이어진 배관」이 하나라도 생기면 잡아낸다.
    #   join_all 이 헤드를 skip 하므로 정상 시 0. 성적 키 이음3·헤드경유.
    ehead = {(min(u, v), max(u, v)) for (k, u, v) in j2 if k == "헤드"}
    head_arms2 = sum(1 for si, a in enumerate(arms)
                     if spots[si]["k"] == "헤드" and len(a) >= 2)
    n_tail = sum(1 for n, vs in _adj(edges).items()
                 if len(vs) == 1 and n in hspot)
    return dict(pts=pts, edges1=e1, edges=edges, spots=spots, arms=arms,
                node_spots=node_spots, hcov=hcov, hnodes=hnodes, hspot=hspot,
                head_kinds=head_kinds,
                j2=j2, j3=j3, passed=passed,
                j4=j4, ring4=ring4, side4=side4,
                j5=j5, ring5=ring5, ups=ups, n_up=len(ups),
                n_up_split=n_up_split,
                n_thru=n_thru, m1=mlen(g.pts, frozenset(g.edges)),
                head_arms2=head_arms2, head_joins=len(ehead),
                n_tail=n_tail, user_sources=user_sources)


def _adj(edges):
    a = defaultdict(set)
    for i, j in edges:
        a[i].add(j)
        a[j].add(i)
    return a


def bodies(pts, edges, hnodes, head_spot_nodes):
    """물덩이 = 이어진 덩어리. 큰 것부터, 헤드 수를 달아서."""
    par = {}

    def find(x):
        par.setdefault(x, x)
        while par[x] != x:
            par[x] = par[par[x]]
            x = par[x]
        return x

    for i, j in edges:
        ri, rj = find(i), find(j)
        if ri != rj:
            par[ri] = rj
    bag = defaultdict(list)
    for i, j in edges:
        bag[find(i)].append((i, j))
    out = []
    for root, es in bag.items():
        ns = {n for e in es for n in e}
        nh = sum(1 for s in hnodes if s & ns)
        out.append(dict(m=mlen(pts, es), n=len(ns), heads=nh, nodes=ns,
                        edges=es))
    out.sort(key=lambda d: -d["m"])
    return out


def split_gaps(pts, edges, hnodes, hspot):
    """★유저가 이을 자리 — 헤드가 든 물덩이끼리 «가장 가까운 곳» [오너 2026-08-07].

    오너: "로직으로 이어 놓고 한 번에 유저에게 안 이은 곳 연결하기로 하자.
           유저가 안 이어진 배관 2개를 클릭하도록 하면 될 것 같은데."

    한 덩이 «안»에서 끊긴 점은 이어도 아무 변화가 없다 — 보여 주지 않는다.
    덩이와 덩이가 갈라진 자리만 낸다. 답 하나가 덩이 둘을 붙이므로
    «헤드 든 덩이 − 1» 이 유저가 클릭할 횟수다.
    반환: (헤드 든 덩이들, [{bi, m, heads, d, at, to, near}])
    """
    bs = bodies(pts, edges, hnodes, hspot)
    hb = [b for b in bs if b["heads"]]
    own = {n: bi for bi, b in enumerate(hb) for n in b["nodes"]}
    grid = defaultdict(list)
    for n in own:
        gput(grid, 2000.0, pts[n][0], pts[n][1], n)
    out = []
    for bi, b in enumerate(hb):
        if bi == 0:
            continue                       # 가장 큰 덩이 = 본망으로 본다
        best = (1e18, None, None)
        for n in b["nodes"]:
            x, y = pts[n]
            for m in set(gnear(grid, 2000.0, x, y, rings=3)):
                if own[m] == bi:
                    continue
                d = math.hypot(pts[m][0] - x, pts[m][1] - y)
                if d < best[0]:
                    best = (d, n, m)
        if best[1] is None:
            continue
        out.append(dict(bi=bi, m=b["m"], heads=b["heads"], d=best[0],
                        at=pts[best[1]], to=pts[best[2]],
                        near=own[best[2]], u=best[1], v=best[2]))
    out.sort(key=lambda r: r["d"])
    return hb, out


def measure(pts, r, hnodes):
    wet = {(i, j) for (i, j) in r["edges"]
           if i in r["reach"] and j in r["reach"]}
    return mlen(pts, wet), sum(1 for s in hnodes if s & r["reach"]), wet


# ★「막힌 자리마다 후보를 줄 세워 오너께 묻는다」는 기계를 걷어냈다
#   [2026-08-07 오너 지시]. 있던 것: out_dir · candidates · draw_ask 와
#   run() 의 «후보마다 이으면 몇 m 더 젖나» 계산.
#   오너: "관이 끊긴 것을 물어보지 말자. 로직으로 이어 놓고 한 번에 유저에게
#          안 이은 곳 연결하기로 해결하자. 유저가 안 이어진 배관 2개를
#          클릭하도록 하면 될 것 같은데."
#   프로그램이 후보를 추천하는 것은 추측이다. 지금은 `split_gaps()` 로
#   «덩이와 덩이가 갈라진 자리»만 내고 고르는 일은 유저가 한다.

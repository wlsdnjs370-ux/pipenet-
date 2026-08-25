# -*- coding: utf-8 -*-
"""[G3] 관경 결정 — 혼합 규칙(지시서 D2).

    nfpc_min = 별표1('가'칸)[담당 헤드 수]
    text     = 선분에서 수직거리 ≤ 1500 mm 인 가장 가까운 치수 텍스트
    dia      = nfpc_min      (text 없음)          → source "nfpc_fallback"
             = nfpc_min      (text < nfpc_min)    → source "nfpc_min"   (안전측)
             = text          (그 외)              → source "text"

모듈 A(`remote30_prototype.py`)의 `_nfpc_min_bore_mm` · `_pipe_diameter` ·
`_extract_dia_text_points` 가 이미 이 규칙 그대로다. **로직을 고치지 않고 옮겼다.**
A 에서는 `build_input_tables` 안의 지역 함수라 그대로 import 할 수 없어 옮겨 적는다.

★단위·좌표(§T1) — 매칭은 **평면 mm** 에서 한다. 전개 결과는 m 이고 한 간선이
여러 배관으로 쪼개지거나 병합되므로, `edge_ref`(kfp 배관 → 원 board 간선)로
되짚어 그 선분에 매칭한다. 전개된 m 좌표로 매칭하면 안 된다.

★담당 헤드 수(§T4) — `worst["loads"][(i,j)]` 를 그대로 쓴다. corridor 안에서 그
간선이 책임지는 «선정된 K개 중의 수» 다. 전체망 하류 헤드 수를 넣으면 관경이
과대해진다.
"""
from __future__ import annotations

import math
import re

# ── 치수 텍스트 패턴 5종 (모듈 A 와 동일) ────────────────────────────────
DIA_PATTERNS = [
    re.compile(r"\b(\d{2,3})\s*A\b"),                      # 25A
    re.compile(r"^\s*(\d{2,3})\s*$"),                      # 순수 숫자
    re.compile(r"[Øø]\s*(\d{2,3})"),                       # Ø25
    re.compile(r"DN\s*(\d{2,3})"),                         # DN25
    re.compile(r"(?<![0-9])(\d{2,3})\s*mm(?![0-9])"),      # 25mm
]
NOISE_KEYWORDS = ("호스", "방수구", "소화전", "옥내", "HOSE", "EA", "KG", "℃",
                  "SET", "SCALE", "PUMP", "펌프", "TANK", "탱크")
VALID_DIA = {15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300}
# 호칭경 텍스트는 보통 배관에 1.5 m 이내로 붙어 있다.
DIA_RANGE_LIMIT_MM = 1500.0


def nfpc_min_bore_mm(head_count: int) -> int:
    """NFPC 103 별표 1 '가' 칸 (폐쇄형 SP) — 담당 헤드 수 → 최소 호칭경."""
    if head_count <= 2:
        return 25
    if head_count <= 3:
        return 32
    if head_count <= 5:
        return 40
    if head_count <= 10:
        return 50
    if head_count <= 30:
        return 65
    if head_count <= 60:
        return 80
    if head_count <= 80:
        return 90
    if head_count <= 100:
        return 100
    if head_count <= 160:
        return 125
    return 150


def extract_dia_text_points(texts) -> list[tuple[float, float, int]]:
    """`world.texts` → [(x, y, dia_mm)]. 노이즈 워드는 버린다.

    입력은 stage1 World 의 `texts`(layer, color, x, y, h, s) 목록이다. 모듈 A 는
    DXF 엔티티 dict(`{"t":"T","v":…,"p":[x,y]}`)를 받으므로 여기가 그 어댑터다.
    DXF 를 다시 읽지 않는다 — 이미 handoff 캐시에 보존돼 있다.
    """
    out: list[tuple[float, float, int]] = []
    for row in texts or ():
        try:
            _lay, _col, x, y, _h, s = row
        except (TypeError, ValueError):
            continue
        v = (s or "").strip()
        if not v:
            continue
        if any(nw in v for nw in NOISE_KEYWORDS):
            continue          # 옥내소화전·헤드 라벨·스펙 표 등
        for pat in DIA_PATTERNS:
            m = pat.search(v)
            if not m:
                continue
            try:
                d = int(m.group(1))
            except ValueError:
                continue
            if d in VALID_DIA:
                out.append((float(x), float(y), d))
            break
    return out


def _point_seg_dist(px, py, ax, ay, bx, by) -> float:
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    if L2 < 1e-9:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def match_diameter_for_segment(a, b, dia_text_pts,
                               limit_mm: float = DIA_RANGE_LIMIT_MM):
    """선분 (a,b) 에 가장 가까운 치수 텍스트 값. 없으면 None. 좌표는 mm."""
    best, best_d = None, float(limit_mm)
    for tx, ty, dia in dia_text_pts:
        d = _point_seg_dist(tx, ty, a[0], a[1], b[0], b[1])
        if d < best_d:
            best_d, best = d, dia
    return best


def decide_bores(net, edge_ref, loads, dia_text_pts, *, pts=None,
                 tree_loads=None) -> dict:
    """kfp 배관마다 (호칭경 mm, 근거). 지시서 §1 공개 시그니처.

    `net`  : 제한 전개 결과 kfp dict (`pipe_data` 를 쓴다)
    `edge_ref` : {pipe_id: (board_i, board_j)}  — §T1 의 역참조
    `loads`    : {(i,j): 담당 헤드 수}          — worst["loads"] 그대로(§T4)
    `dia_text_pts` : [(x, y, dia_mm)] — `extract_dia_text_points` 결과
    `pts`      : board 노드 좌표(mm). 없으면 텍스트 매칭을 건너뛴다.
    `tree_loads` : {pipe_id: 담당 헤드 수} — 역참조가 **없는** 배관용.
        헤드 접속관·가지 상승은 도면에 그려진 선이 아니라 대응할 board 간선이
        없다. 그 자리를 0 으로 두면 별표1 이 전부 25A 를 주고, 제 아래 헤드
        스무 개를 받는 가지 상승관까지 25A 가 된다. 망에서 직접 센 값을 쓴다.
        **안 넘기면 여기서 직접 센다** — 부르는 쪽이 잊으면 조용히 25A 가 되고,
        그 잘못은 표를 한참 들여다봐야 보인다(실측: 검사 경로만 그랬다).

    반환: {pipe_id: (dia_mm, source)} · source ∈ {text, nfpc_min, nfpc_fallback}
    """
    pipes = (net or {}).get("pipe_data") or {}
    if tree_loads is None and any(pid not in edge_ref for pid in pipes):
        from services.cad_import.design.restrict import tree_loads as _tl
        tree_loads = _tl(net)
    out: dict = {}
    for pid in pipes:
        ref = edge_ref.get(pid)
        n_head = 0
        if ref is not None:
            i, j = ref
            n_head = int(loads.get((min(i, j), max(i, j)), 0))
        elif tree_loads:
            n_head = int(tree_loads.get(pid, 0))
        nfpc_min = nfpc_min_bore_mm(n_head)

        text = None
        if ref is not None and pts is not None:
            i, j = ref
            if 0 <= i < len(pts) and 0 <= j < len(pts):
                text = match_diameter_for_segment(pts[i], pts[j], dia_text_pts)

        if text is None:
            out[pid] = (nfpc_min, "nfpc_fallback")
        elif text < nfpc_min:
            # 안전측 — 도면 치수가 별표1 최소보다 작으면 별표1 을 따른다.
            out[pid] = (nfpc_min, "nfpc_min")
        else:
            out[pid] = (text, "text")
    return out


def source_counts(bores: dict) -> dict:
    """근거 집계 — 화면·meta 에 남긴다. text 가 0% 면 어댑터가 죽은 것이다."""
    counts = {"text": 0, "nfpc_min": 0, "nfpc_fallback": 0}
    for _dia, src in (bores or {}).values():
        if src in counts:
            counts[src] += 1
    return counts

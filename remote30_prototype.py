"""Remote 30 프로토타입 — DXF 한 장 → 4-stage 파이프라인 (배관망 → 30 헤드 → CSV/XLSX → SDF).

각 stage 는 JSON 직렬화 가능한 진행 이벤트를 yield 한다. 호출자(서버) 가 이걸 SSE 로 클라이언트에 흘림.

Stages
------
0  parse_dxf            : ezdxf 로 modelspace 읽어 raw entity + 레이어 정보
1  pipenet_only_filter  : "배관망만" 카테고리(PIPE/HEAD/TEXT/0/L1-4) 만 통과 + CAD hidden 차단
2  select_top30_heads   : G₀ 그래프 빌드 → 알람밸브 자동 식별 → 가장 불리한 K개 헤드 + subgraph
3  build_input_tables   : Nodes/Pipes/Nozzles/Fittings/Equipment 5 테이블 + Meta 시트
4  emit_sdf             : PIPENET .sdf XML emit (Project > Network-spray > ...)

진행 이벤트 형식
----------------
{"type": "stage", "stage": 1, "label": "...", "status": "running"|"done", "elapsed_ms": ...}
{"type": "entities", "stage": 0|1|2, "entities": [...], "bbox": {...}, "layers": [...]}
{"type": "tables_preview", "tables": {...}}
{"type": "done", "outputs": {"xlsx": "...", "csv_*": "...", "sdf": "..."}}
{"type": "error", "stage": ..., "message": "..."}
"""

from __future__ import annotations

# ── core/ 라이브러리 경로 (repo 정리: 루트 라이브러리 → core/ 이동) ──
import sys as _sys
from pathlib import Path as _Path
_CORE = _Path(__file__).resolve().parent / "core"
if _CORE.is_dir() and str(_CORE) not in _sys.path:
    _sys.path.insert(0, str(_CORE))

import csv
import gzip
import hashlib
import heapq
import math
import os
import pickle
import re
import time
import warnings
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator
from xml.dom import minidom

import ezdxf
from ezdxf.math import Matrix44, Vec3


# ────────────────────────────────────────────────────────────────────────────
# 자산 파일 경로 해석 — 환경변수 → 모듈 디렉토리 fallback
# ────────────────────────────────────────────────────────────────────────────
# emit_sdf 는 두 가지 자산 파일에 의존한다:
#   1. Template SDF — PIPENET 의 Graphics 블록(아이소매트릭 표시 메타) 보존용
#   2. Standard SLF — 6 schedule 정의 + 표준 노즐/펌프 라이브러리, 결과 폴더에 동봉
#
# 본 모듈을 Linux 서버 / 다른 PC / Docker / CI 등 다양한 환경에서 실행 가능하게 하기 위해
# 절대 경로 하드코딩 대신 다음 우선순위로 해석한다:
#   ① 환경변수 (REMOTE30_TEMPLATE_SDF, REMOTE30_STANDARD_SLF)
#   ② 모듈 디렉토리 (=`__file__` 의 부모) 기준 상대 파일명
#
# 두 단계 모두 실패하면 명확한 RuntimeWarning 을 발행하고 None 을 반환.
# 호출 측은 None 을 받아 fallback 경로(template 없이 빈 SDF / SLF 동봉 생략)를
# 택할 수 있지만, 그 영향(아이소매트릭 누락 / diameter Unset)을 사용자가 인지하게 된다.

_MODULE_DIR = Path(__file__).resolve().parent
TEMPLATE_SDF_FILENAME = "3-1형_자연낙차_LSP_4F_OA_지하층포함_120m~200m미만_6.6K로 감압_알람밸브.sdf"
STANDARD_SLF_FILENAME = "2. Pipenet_hand_FX28.slf"


def _resolve_asset(env_var: str, default_filename: str, *, role: str) -> Path | None:
    """환경변수 → 모듈 디렉토리 순으로 자산 파일을 찾는다. 못 찾으면 None.

    Args:
        env_var: 절대/상대 경로를 담은 환경변수 이름. 비어있으면 모듈 디렉토리로 폴백.
        default_filename: 모듈 디렉토리에서 찾을 파일명.
        role: 경고 메시지에 쓰일 자산 역할 설명 ("Template SDF" 등).

    Returns:
        해석된 절대 경로 (Path), 또는 None (둘 다 실패).
    """
    env_val = os.environ.get(env_var, "").strip()
    if env_val:
        candidate = Path(env_val).expanduser()
        if not candidate.is_absolute():
            candidate = (Path.cwd() / candidate).resolve()
        else:
            candidate = candidate.resolve()
        if candidate.is_file():
            return candidate
        warnings.warn(
            f"[remote30_prototype] 환경변수 {env_var}='{env_val}' 지정 — "
            f"하지만 '{candidate}' 에 {role} 파일이 없음. 모듈 디렉토리 fallback 시도.",
            RuntimeWarning, stacklevel=3,
        )

    candidate = (_MODULE_DIR / default_filename).resolve()
    if candidate.is_file():
        return candidate
    return None


def resolve_template_sdf() -> Path | None:
    """PIPENET Graphics 블록(아이소매트릭 메타) 보존용 template SDF 경로."""
    return _resolve_asset(
        "REMOTE30_TEMPLATE_SDF", TEMPLATE_SDF_FILENAME, role="Template SDF",
    )


def resolve_standard_slf() -> Path | None:
    """6 schedule + 표준 노즐/펌프 정의가 담긴 표준 SLF 경로."""
    return _resolve_asset(
        "REMOTE30_STANDARD_SLF", STANDARD_SLF_FILENAME, role="Standard SLF",
    )

# Optional — sprinkler_remote30_extractor 의 layer 카테고리 분류 활용
try:
    from sprinkler_remote30_extractor import Remote30Settings, layer_match
except ImportError:
    Remote30Settings = None  # type: ignore
    layer_match = None  # type: ignore


# ────────────────────────────────────────────────────────────────────────────
# 0) ezdxf modelspace 파싱 + 매트릭스 보정 + hidden 차단 → 캔버스용 entity
# ────────────────────────────────────────────────────────────────────────────

from remote30_constants import (PIPENET_CATEGORIES, KEEP_BASE_LAYERS, _DIA_TEXT_PATTERNS, _DIA_TEXT_NOISE_KW, _VALID_DIA_MM, _FLOOR_LABEL_PATTERNS, _FLOOR_LABEL_SPECIAL, MACHINE_ROOM_SP_LAYERS, SNAP_TOL_MM, HEAD_BRIDGE_MAX_MM, SOURCE_BRIDGE_MAX_MM, ANCHOR_W_MARGIN_MM, MIN_PIPE_EDGE_MM, TEE_SPLIT_MAX_MM, CLOSED_PL_TOL_MM, LADDER_MAX_RUNG_MM, LADDER_MIN_RAIL_RATIO, LADDER_PARALLEL_COS, LADDER_MAX_ITER, STEEL_PIPE_TYPE, STEEL_C_FACTOR, CPVC_PIPE_TYPE, CPVC_C_FACTOR, ORTHO_SNAP_TOL_DEG, FX_SPEC_PROFILES, FX_DEFAULT_PROFILE, AV_EQ_LEN_M, FX_SCHEDULE_ROUGHNESS, FX_RISE_M, fx_schedule_name, fx_geometry_key)  # noqa: E501  (Phase2b core)

# ── 초대형 XREF 도면 예산 가드 (기계실/계통도 파싱) ──────────────────────────
# 142MB LH 지하층배관도는 최상위 INSERT 61개를 폭발하면 leaf 597k개가 되는데
# 그중 배관망 추출에 실제 쓰는 PIPE+HEAD 는 ~6.7k(1%)뿐이고 나머지 99% 는 배경.
# 도면 전체가 PIPE 레이어 위의 단일 INSERT(소방배관) 안에 nested 돼 있어 미리보기
# 경로의 '최상위 카테고리별 배경 분리' 가드로는 안 걸린다(그 블록이 PIPE로 분류되어
# 통째 폭발). 그래서 폭발 도중 **leaf 단위**로 카테고리를 컷한다.
#
# 실측(2026-07-07, _bg_probe): ARCH/EXCLUDE 만 스킵하면 시간 이득 0 — ARCH 280k는
# XREF 단순 LINE 이라 렌더가 거의 공짜다. 파싱 비용의 전부(112→37초, 75초)는 OTHER
# 카테고리(LINE 173k + PL 111k + ARC 21k)에 있다. 즉 실질 단축은 OTHER 컷이 필수.
# 대가: OTHER 의 line/PL 은 형태상 배관이 될 수도 있어(레이어명 불신 원칙,
# [[lh-basement-fire-layer]]) 이론상 '키워드 미매칭 배관 레이어'가 숨을 수 있다. 단
# (1) 예산(120k leaf) 초과하는 초대형 XREF 도면에서만 발동, (2) 배관 키워드 목록이
# 넓고(SP/배관/소방/PIPE/PIPING/FIRE/HYD/가지관/FLEX), (3) 숨겨져도 추출이 조용히
# 틀리는 게 아니라 '망 없음'으로 드러난다. 사용자 승인 하에 OTHER 포함(2026-07-07).
BG_ENTITY_BUDGET = 120_000  # 폭발 예상 leaf 이 이 값을 넘으면 배경 leaf 스킵 발동
_BG_SKIP_CATEGORIES = frozenset({"ARCH", "EXCLUDE", "OTHER"})


def _categorize_layer(name: str) -> str:
    """Remote30Settings 기준 layer 카테고리. 가능하면 외부 모듈 사용.

    ALARM 카테고리 추가 (2026-06-08) — 알람밸브 키워드 매칭되는 레이어
    (예: "RISER", "라이저") 를 별도 분류해서 filter_pipenet_only 통과시킴.
    이전엔 OTHER 로 떨어져서 _find_source 가 RISER 의 INSERT 를 볼 수 없었음.
    """
    if Remote30Settings is None or layer_match is None:
        # fallback heuristic
        n = name.lower()
        if any(k in n for k in ("소화기", "옥내소화전", "자동식", "co2")):
            return "EXCLUDE"
        if any(k in name for k in ("HEAD", "헤드", "SP-H", "하향식", "상향식", "헤드반경")):
            return "HEAD"
        if any(k in name.upper() for k in ("ALARM", "RISER", "라이저", "STAND-PIPE")):
            return "ALARM"
        if any(k in name for k in ("SP", "배관", "소방", "가지관", "후렉시블", "FLEX")) \
                or any(k in name.upper() for k in ("PIPE", "PIPING")):
            return "PIPE"
        if any(k in n for k in ("text", "문자")) or "TEX" in name:
            return "TEXT"
        if any(k in name for k in ("벽", "건축", "WALL", "ARCH", "DIM", "SHEET", "AREA")):
            return "ARCH"
        return "OTHER"
    s = Remote30Settings()
    # 콘텐츠(HEAD/ALARM/PIPE/TEXT) 신호가 ARCH 를 이긴다. EXCLUDE 만 최우선.
    # arch-first 였을 때 "SHEET-TEXT" 등 콘텐츠 레이어가 건축으로 흡수되던 오류 방지.
    if layer_match(name, s.exclude_layer_keywords):
        return "EXCLUDE"
    if layer_match(name, s.head_layer_keywords):
        return "HEAD"
    # ALARM 검사를 PIPE 보다 먼저 — "RISER" 가 "SP" 와 겹치지 않지만 우선순위 명시
    if layer_match(name, s.alarm_valve_keywords):
        return "ALARM"
    if layer_match(name, s.pipe_layer_keywords):
        return "PIPE"
    if layer_match(name, s.text_layer_keywords):
        return "TEXT"
    if layer_match(name, s.arch_layer_keywords):
        return "ARCH"
    return "OTHER"


class _BBoxAccum:
    """좌표 누적 후 percentile-based robust bbox 계산.

    raw min/max 는 outlier 한 점이 bbox 폭주시켜 캔버스 fit 시 도면이 매우 작게
    보이는 문제 발생 (WIPEOUT 의 (1e30, 1e30), 잘못 변환된 nested INSERT 좌표,
    paper space 잔재 등). percentile [pct_low, pct_high] 으로 main cluster bbox
    를 구해 안정적인 초기 시야를 제공.

    좌표 단위 percentile (entity 단위 아님) — PL N vertex 는 N 점으로 계산.
    """

    __slots__ = ("xs", "ys")

    def __init__(self) -> None:
        self.xs: list[float] = []
        self.ys: list[float] = []

    def add(self, x: float, y: float) -> None:
        # NaN / inf 즉시 거부 (DXF 파싱 에러로 가끔 발생)
        if x != x or y != y or x in (float("inf"), float("-inf")) or y in (float("inf"), float("-inf")):
            return
        self.xs.append(x)
        self.ys.append(y)

    def finalize(
        self,
        pct_low: float = 0.5,
        pct_high: float = 99.5,
        margin_ratio: float = 0.02,
        min_margin: float = 50.0,
    ) -> list[float]:
        """robust bbox [xmin, ymin, xmax, ymax]. 좌표 없으면 [0,0,1,1] fallback."""
        if not self.xs:
            return [0.0, 0.0, 1.0, 1.0]
        n = len(self.xs)
        xs = sorted(self.xs)
        ys = sorted(self.ys)
        lo = max(int(n * pct_low / 100.0), 0)
        hi = min(int(n * pct_high / 100.0), n - 1)
        if hi <= lo:
            hi = lo
        x_min, x_max = xs[lo], xs[hi]
        y_min, y_max = ys[lo], ys[hi]
        w = x_max - x_min
        h = y_max - y_min
        mx = max(w * margin_ratio, min_margin)
        my = max(h * margin_ratio, min_margin)
        return [x_min - mx, y_min - my, x_max + mx, y_max + my]

    def outlier_stats(self, pct_low: float = 0.5, pct_high: float = 99.5) -> dict:
        """raw vs robust bbox 비교 + outlier 점 수. 진단용."""
        if not self.xs:
            return {"coord_count": 0, "outlier_points": 0,
                    "raw_bbox": [0, 0, 0, 0], "robust_bbox": [0, 0, 1, 1],
                    "bbox_ratio": 1.0}
        n = len(self.xs)
        xs = sorted(self.xs); ys = sorted(self.ys)
        lo = max(int(n * pct_low / 100.0), 0)
        hi = min(int(n * pct_high / 100.0), n - 1)
        outliers = lo + (n - 1 - hi)
        raw_bbox = [xs[0], ys[0], xs[-1], ys[-1]]
        rob_bbox = [xs[lo], ys[lo], xs[hi], ys[hi]]
        raw_w = max(raw_bbox[2] - raw_bbox[0], 1.0)
        raw_h = max(raw_bbox[3] - raw_bbox[1], 1.0)
        rob_w = max(rob_bbox[2] - rob_bbox[0], 1.0)
        rob_h = max(rob_bbox[3] - rob_bbox[1], 1.0)
        # raw bbox 가 robust 보다 N 배 크면 outlier 의심 — 화면 fit 시 도면이 1/N 로 보임
        bbox_ratio = max(raw_w / rob_w, raw_h / rob_h)
        return {
            "coord_count": n,
            "outlier_points": outliers * 2,  # x + y 양쪽
            "raw_bbox": raw_bbox,
            "robust_bbox": rob_bbox,
            "bbox_ratio": round(bbox_ratio, 2),
        }


@dataclass(slots=True)
class ParsedDxfBundle:
    """Stage 0 출력 — 캔버스가 직접 그릴 수 있는 entity dict + 메타."""

    entities: list[dict] = field(default_factory=list)
    bbox: list[float] = field(default_factory=lambda: [0.0, 0.0, 1.0, 1.0])
    layers: list[dict] = field(default_factory=list)
    hidden_layers: set[str] = field(default_factory=set)
    layer_visibility: dict[str, dict] = field(default_factory=dict)
    # entity index → source meta (graph stage 에서 좌표→layer 매칭에 사용)
    layer_counts: dict[str, int] = field(default_factory=dict)
    # robust bbox 진단 (outlier 가 있을 때 디버깅 + 라벨에 표시)
    bbox_diagnostics: dict = field(default_factory=dict)


def _insert_matrix(insert_entity) -> Matrix44:
    """AutoCAD 표준 INSERT 변환 매트릭스 — M·local = world."""
    ix = float(insert_entity.dxf.insert.x)
    iy = float(insert_entity.dxf.insert.y)
    try:
        iz = float(insert_entity.dxf.insert.z)
    except Exception:
        iz = 0.0
    sx = float(getattr(insert_entity.dxf, "xscale", 1.0) or 1.0)
    sy = float(getattr(insert_entity.dxf, "yscale", 1.0) or 1.0)
    sz = float(getattr(insert_entity.dxf, "zscale", 1.0) or 1.0)
    rot_rad = math.radians(float(getattr(insert_entity.dxf, "rotation", 0.0) or 0.0))
    block = insert_entity.doc.blocks.get(insert_entity.dxf.name) if insert_entity.doc else None
    if block is not None:
        try:
            bx = float(block.base_point.x)
            by = float(block.base_point.y)
            bz = float(block.base_point.z) if hasattr(block.base_point, "z") else 0.0
        except Exception:
            bx = by = bz = 0.0
    else:
        bx = by = bz = 0.0
    return Matrix44.chain(
        Matrix44.translate(-bx, -by, -bz),
        Matrix44.scale(sx, sy, sz),
        Matrix44.z_rotate(rot_rad),
        Matrix44.translate(ix, iy, iz),
    )


def _t(matrix: Matrix44 | None, x: float, y: float) -> tuple[float, float]:
    if matrix is None:
        return float(x), float(y)
    v = matrix.transform(Vec3(float(x), float(y), 0.0))
    return float(v.x), float(v.y)


def _uniform_scale(matrix: Matrix44 | None) -> float:
    """matrix 의 등방(uniform) 스케일 추정 — 단위 X 벡터를 변환한 길이.

    ARC/CIRCLE 반지름을 INSERT 변환 후에도 맞추기 위함. matrix None 이면 1.0.
    """
    if matrix is None:
        return 1.0
    p0 = matrix.transform(Vec3(0.0, 0.0, 0.0))
    p1 = matrix.transform(Vec3(1.0, 0.0, 0.0))
    return math.hypot(p1.x - p0.x, p1.y - p0.y)


def _hatch_outline_paths(entity, matrix: Matrix44 | None) -> list[list[list[float]]]:
    """HATCH 경계 폴리라인 추출 — PolylinePath vertices 우선, 없으면 EdgePath(Line/Arc).

    각 path 의 [x, y] 점목록(연속 중복점 제거)을 모아 반환하며 빈 path 는 제외한다.
    호출부에서 bbox 갱신 + 최대 path 선택을 수행한다.
    """
    paths_out: list[list[list[float]]] = []
    for path in entity.paths:
        pts: list[list[float]] = []
        for vertex in getattr(path, "vertices", []) or []:
            try:
                x, y = _t(matrix, vertex[0], vertex[1])
                pts.append([x, y])
            except Exception:
                continue
        if not pts:
            for edge in getattr(path, "edges", []) or []:
                et = type(edge).__name__
                try:
                    if et == "LineEdge":
                        x1, y1 = _t(matrix, edge.start[0], edge.start[1])
                        x2, y2 = _t(matrix, edge.end[0], edge.end[1])
                        pts.append([x1, y1]); pts.append([x2, y2])
                    elif et == "ArcEdge":
                        cx = float(edge.center[0]); cy = float(edge.center[1])
                        r = float(edge.radius)
                        sa = float(edge.start_angle); ea = float(edge.end_angle)
                        if ea < sa:
                            ea += 360.0
                        for k in range(9):
                            ang = math.radians(sa + (ea - sa) * k / 8)
                            x, y = _t(matrix, cx + r * math.cos(ang), cy + r * math.sin(ang))
                            pts.append([x, y])
                except Exception:
                    continue
        if len(pts) > 1:
            pts = [pts[0]] + [p for prev, p in zip(pts, pts[1:]) if p != prev]
        if pts:
            paths_out.append(pts)
    return paths_out


def parse_dxf_bundle(dxf_path: Path) -> ParsedDxfBundle:
    """ezdxf 로 modelspace 파싱 → 캔버스용 entity dict 리스트.

    레이어 hidden 차단(is_off/is_frozen/color<0) + INSERT mirror 매트릭스 적용.
    """
    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()
    bundle = ParsedDxfBundle()

    # 레이어 가시성
    for ly in doc.layers:
        try:
            color = int(ly.dxf.color)
        except Exception:
            color = 7
        name = str(ly.dxf.name)
        is_off = bool(ly.is_off())
        is_frozen = bool(ly.is_frozen())
        try:
            no_plot = int(getattr(ly.dxf, "plot", 1)) == 0
        except Exception:
            no_plot = False
        bundle.layer_visibility[name] = {
            "is_off": is_off,
            "is_frozen": is_frozen,
            "color": color,
            "no_plot": no_plot,
        }
        # CAD 화면에 안 보이는 것만 숨김 (off/frozen/color<0). plot=0(비출력) 레이어는
        # 화면엔 보이므로 렌더해 실제 도면과 동일 규격 유지.
        if is_off or is_frozen or color < 0:
            bundle.hidden_layers.add(name)

    bbox_acc = _BBoxAccum()

    def _upd(x: float, y: float) -> None:
        bbox_acc.add(x, y)

    MAX_DEPTH = 10

    def _render(e, matrix=None, layer_override=None, depth=0):
        etype = e.dxftype()
        own = getattr(e.dxf, "layer", "")
        if layer_override is not None and own in ("0", ""):
            layer = layer_override
        else:
            layer = own or (layer_override or "")
        if layer in bundle.hidden_layers:
            return
        if int(getattr(e.dxf, "invisible", 0) or 0) == 1:
            return
        try:
            if etype == "LINE":
                x1, y1 = _t(matrix, e.dxf.start.x, e.dxf.start.y)
                x2, y2 = _t(matrix, e.dxf.end.x, e.dxf.end.y)
                bundle.entities.append({"t": "L", "l": layer, "p": [x1, y1, x2, y2]})
                _upd(x1, y1); _upd(x2, y2)
            elif etype == "ARC":
                cx, cy = _t(matrix, e.dxf.center.x, e.dxf.center.y)
                r = float(e.dxf.radius) * _uniform_scale(matrix)
                bundle.entities.append({"t": "A", "l": layer, "c": [cx, cy], "r": r,
                                       "a": [float(e.dxf.start_angle), float(e.dxf.end_angle)]})
                _upd(cx - r, cy - r); _upd(cx + r, cy + r)
            elif etype == "CIRCLE":
                cx, cy = _t(matrix, e.dxf.center.x, e.dxf.center.y)
                r = float(e.dxf.radius) * _uniform_scale(matrix)
                bundle.entities.append({"t": "C", "l": layer, "c": [cx, cy], "r": r})
                _upd(cx - r, cy - r); _upd(cx + r, cy + r)
            elif etype == "LWPOLYLINE":
                pts = [list(_t(matrix, p[0], p[1])) for p in e.get_points()]
                if pts:
                    for x, y in pts:
                        _upd(x, y)
                    bundle.entities.append({"t": "PL", "l": layer, "p": pts})
            elif etype == "POLYLINE":
                pts = [list(_t(matrix, v.dxf.location.x, v.dxf.location.y)) for v in e.vertices]
                if pts:
                    for x, y in pts:
                        _upd(x, y)
                    bundle.entities.append({"t": "PL", "l": layer, "p": pts})
            elif etype == "INSERT":
                ix_w, iy_w = _t(matrix, e.dxf.insert.x, e.dxf.insert.y)
                if depth == 0:
                    bundle.entities.append({"t": "I", "l": layer, "p": [ix_w, iy_w],
                                           "n": str(e.dxf.name)})
                _upd(ix_w, iy_w)
                # ARCH/EXCLUDE 레이어 블록도 폭발 — 실제 CAD 도면과 동일하게 건축 배경 렌더.
                if depth >= MAX_DEPTH:
                    return
                try:
                    my_m = _insert_matrix(e)
                except Exception:
                    my_m = None
                if matrix is not None and my_m is not None:
                    combined = Matrix44.chain(my_m, matrix)
                elif my_m is not None:
                    combined = my_m
                else:
                    combined = matrix
                block = e.doc.blocks.get(e.dxf.name) if e.doc else None
                if block is not None:
                    for child in block:
                        _render(child, matrix=combined, layer_override=layer, depth=depth + 1)
            elif etype == "TEXT":
                x, y = _t(matrix, e.dxf.insert.x, e.dxf.insert.y)
                raw = str(e.dxf.text)[:60]
                bundle.entities.append({"t": "T", "l": layer, "p": [x, y], "v": raw})
                _upd(x, y)
            elif etype in ("MTEXT", "ATTRIB", "ATTDEF"):
                x, y = _t(matrix, e.dxf.insert.x, e.dxf.insert.y)
                raw = str(getattr(e, "text", "") or getattr(e.dxf, "text", ""))[:60]
                if raw:
                    bundle.entities.append({"t": "T", "l": layer, "p": [x, y], "v": raw})
                _upd(x, y)
            elif etype == "SPLINE":
                try:
                    pts = [list(_t(matrix, pt[0], pt[1])) for pt in e.flattening(1.0)]
                except Exception:
                    pts = []
                if pts:
                    for x, y in pts:
                        _upd(x, y)
                    bundle.entities.append({"t": "PL", "l": layer, "p": pts})
            elif etype == "ELLIPSE":
                try:
                    pts = [list(_t(matrix, pt[0], pt[1])) for pt in e.flattening(0.5)]
                except Exception:
                    pts = []
                if pts:
                    for x, y in pts:
                        _upd(x, y)
                    bundle.entities.append({"t": "PL", "l": layer, "p": pts})
            elif etype == "HATCH":
                paths_out = _hatch_outline_paths(e, matrix)
                for pts in paths_out:
                    for x, y in pts:
                        _upd(x, y)
                if paths_out:
                    biggest = max(paths_out, key=len)
                    bundle.entities.append({"t": "H", "l": layer, "p": biggest})
            elif etype in ("SOLID", "3DFACE", "TRACE"):
                verts = []
                for attr in ("vtx0", "vtx1", "vtx2", "vtx3"):
                    try:
                        v = getattr(e.dxf, attr)
                        x, y = _t(matrix, v.x, v.y)
                        verts.append([x, y])
                    except AttributeError:
                        break
                if len(verts) >= 2 and verts[-1] == verts[-2]:
                    verts.pop()
                if len(verts) >= 3:
                    for x, y in verts:
                        _upd(x, y)
                    bundle.entities.append({"t": "S", "l": layer, "p": verts})
            elif etype == "DIMENSION":
                try:
                    for v in e.virtual_entities():
                        _render(v, matrix=matrix, layer_override=layer)
                except Exception:
                    pass
        except Exception:
            pass

    for e in msp:
        _render(e)

    bundle.bbox = bbox_acc.finalize()
    bundle.bbox_diagnostics = bbox_acc.outlier_stats()

    # 레이어 통계 + 카테고리
    layer_counts: Counter[str] = Counter(en["l"] for en in bundle.entities)
    bundle.layer_counts = dict(layer_counts)
    for name in sorted(layer_counts):
        info = bundle.layer_visibility.get(name, {})
        bundle.layers.append({
            "name": name,
            "count": layer_counts[name],
            "auto_category": _categorize_layer(name),
            "color": info.get("color", 7),
            "is_off": info.get("is_off", False),
            "is_frozen": info.get("is_frozen", False),
            "visible": not (info.get("is_off", False) or info.get("is_frozen", False) or info.get("color", 7) < 0),
        })
    return bundle


_PARSE_CACHE_VERSION = 1
_PARSE_CACHE_DIR = _Path(__file__).resolve().parent / "data" / "parse_cache"


def _file_content_key(dxf_path: Path) -> str:
    h = hashlib.blake2b(digest_size=16)
    with open(dxf_path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def parse_dxf_bundle_cached(dxf_path: Path) -> ParsedDxfBundle:
    """parse_dxf_bundle 을 파일 내용 해시 키로 디스크 캐시해 재파싱을 건너뛴다."""
    try:
        key = _file_content_key(dxf_path)
    except OSError:
        return parse_dxf_bundle(dxf_path)
    cache_path = _PARSE_CACHE_DIR / f"v{_PARSE_CACHE_VERSION}_{key}.pkl.gz"
    if cache_path.is_file():
        try:
            with gzip.open(cache_path, "rb") as f:
                obj = pickle.load(f)
            if isinstance(obj, ParsedDxfBundle):
                return obj
        except Exception:
            pass
    bundle = parse_dxf_bundle(dxf_path)
    try:
        _PARSE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        tmp = cache_path.with_suffix(".gz.tmp")
        with gzip.open(tmp, "wb") as f:
            pickle.dump(bundle, f, protocol=pickle.HIGHEST_PROTOCOL)
        tmp.replace(cache_path)
    except Exception:
        pass
    return bundle


def parse_dxf_for_view(dxf_path: Path, *, include_hidden_layers: bool = True,
                        keep_nested_insert_markers: bool = False,
                        skip_background_over_budget: bool = False) -> dict:
    """계통도 등 '시각화 우선' 용 파싱 — parse_dxf_bundle 의 보강 버전.

    parse_dxf_bundle 과 차이:
        ① include_hidden_layers=True (기본) — is_off/is_frozen/color<0 layer 도 모두 포함
        ② POINT / LEADER / MLEADER / 3DPOLYLINE / RAY / XLINE / WIPEOUT 등 추가 type
        ③ keep_nested_insert_markers — depth>0 의 nested INSERT 도 표지 표시 (옵션)
        ④ skip / error counter 반환 — 어떤 entity type 이 못 그려졌는지 보고

    Args:
        dxf_path: DXF 파일 경로.
        include_hidden_layers: True 면 hidden 무시 (모든 layer 추출).
        keep_nested_insert_markers: True 면 nested INSERT 마커도 entity 로 표시.

    Returns:
        dict {
            "entities": [...],            # parse_dxf_bundle 와 동일 포맷
            "layers": [...],
            "bbox": [xmin, ymin, xmax, ymax],
            "skipped": {etype: count, ...},   # 미지원 / 변환실패 entity 통계
            "total_msp_entities": int,        # modelspace 최상위 entity 수
        }
    """
    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()

    entities: list[dict] = []
    layer_visibility: dict[str, dict] = {}
    hidden_layers: set[str] = set()
    skipped: Counter = Counter()

    # 레이어 가시성 (정보만 — include_hidden_layers=True 면 skip 안 함)
    for ly in doc.layers:
        try:
            color = int(ly.dxf.color)
        except Exception:
            color = 7
        name = str(ly.dxf.name)
        is_off = bool(ly.is_off())
        is_frozen = bool(ly.is_frozen())
        layer_visibility[name] = {"is_off": is_off, "is_frozen": is_frozen, "color": color}
        if is_off or is_frozen or color < 0:
            hidden_layers.add(name)

    bbox_acc = _BBoxAccum()

    def _upd(x: float, y: float) -> None:
        bbox_acc.add(x, y)

    MAX_DEPTH = 12  # 계통도는 nested 깊을 수 있음 — 약간 여유

    # ── 예산 가드: 폭발 예상 leaf 수 추정 후 초과 시 배경(_BG_SKIP_CATEGORIES) leaf 스킵 ──
    # 최상위 INSERT 를 블록정의 단위로만 세어(같은 블록 재INSERT 는 depth별 1회 memo)
    # 렌더 없이 leaf 수를 O(예산×5) 로 유계 추정한다. 미리보기 경로의 _bg_leaf_estimate
    # 와 동일한 알고리즘 — 단 여기선 최상위 분리 없이 전체 합을 기준으로 판정한다
    # (LH 는 배관/배경이 한 최상위 INSERT 안에 섞여 최상위 카테고리로는 못 가름).
    skip_cats: frozenset = frozenset()
    bg_skipped = False
    bg_leaf_estimate = 0
    if skip_background_over_budget:
        _ceiling = BG_ENTITY_BUDGET * 5
        _memo: dict[tuple, int] = {}

        def _block_leaves(block_name, d):
            if d >= MAX_DEPTH or doc is None:
                return 0
            key = (block_name, d)
            if key in _memo:
                return _memo[key]
            blk = doc.blocks.get(block_name)
            total = 0
            if blk is not None:
                for child in blk:
                    if child.dxftype() == "INSERT":
                        total += _block_leaves(child.dxf.name, d + 1)
                    else:
                        total += 1
                    if total >= _ceiling:
                        total = _ceiling
                        break
            _memo[key] = total
            return total

        for _e in msp:
            try:
                if _e.dxftype() == "INSERT":
                    bg_leaf_estimate += _block_leaves(_e.dxf.name, 0)
                else:
                    bg_leaf_estimate += 1
            except Exception:
                continue
            if bg_leaf_estimate >= _ceiling:
                break
        if bg_leaf_estimate > BG_ENTITY_BUDGET:
            skip_cats = _BG_SKIP_CATEGORIES
            bg_skipped = True

    _cat_cache: dict[str, str] = {}

    def _skip_layer(layer_name: str) -> bool:
        cat = _cat_cache.get(layer_name)
        if cat is None:
            cat = _categorize_layer(layer_name)
            _cat_cache[layer_name] = cat
        return cat in skip_cats

    def _render(e, matrix=None, layer_override=None, depth=0):
        etype = e.dxftype()
        own = getattr(e.dxf, "layer", "")
        if layer_override is not None and own in ("0", ""):
            layer = layer_override
        else:
            layer = own or (layer_override or "")
        # ★ hidden 무시 (계통도 모드) 또는 차단 (기본 모드)
        if not include_hidden_layers and layer in hidden_layers:
            return
        if int(getattr(e.dxf, "invisible", 0) or 0) == 1:
            return
        # ★ 예산 초과 시 배경 leaf 스킵 — INSERT 는 하위에 전경이 섞여 있을 수 있어
        #   무조건 재귀하되(마커만 생략), 나머지 leaf 타입만 카테고리로 컷.
        if skip_cats and etype != "INSERT" and _skip_layer(layer):
            skipped["BG_BUDGET_SKIP"] += 1
            return
        try:
            if etype == "LINE":
                x1, y1 = _t(matrix, e.dxf.start.x, e.dxf.start.y)
                x2, y2 = _t(matrix, e.dxf.end.x, e.dxf.end.y)
                entities.append({"t": "L", "l": layer, "p": [x1, y1, x2, y2]})
                _upd(x1, y1); _upd(x2, y2)
            elif etype == "ARC":
                cx, cy = _t(matrix, e.dxf.center.x, e.dxf.center.y)
                r = float(e.dxf.radius) * _uniform_scale(matrix)
                entities.append({"t": "A", "l": layer, "c": [cx, cy], "r": r,
                                  "a": [float(e.dxf.start_angle), float(e.dxf.end_angle)]})
                _upd(cx - r, cy - r); _upd(cx + r, cy + r)
            elif etype == "CIRCLE":
                cx, cy = _t(matrix, e.dxf.center.x, e.dxf.center.y)
                r = float(e.dxf.radius) * _uniform_scale(matrix)
                entities.append({"t": "C", "l": layer, "c": [cx, cy], "r": r})
                _upd(cx - r, cy - r); _upd(cx + r, cy + r)
            elif etype == "LWPOLYLINE":
                pts = [list(_t(matrix, p[0], p[1])) for p in e.get_points()]
                if pts:
                    for x, y in pts: _upd(x, y)
                    entities.append({"t": "PL", "l": layer, "p": pts})
            elif etype == "POLYLINE":
                pts = []
                for v in e.vertices:
                    try:
                        loc = v.dxf.location
                        x, y = _t(matrix, loc.x, loc.y)
                        pts.append([x, y])
                    except Exception:
                        continue
                if pts:
                    for x, y in pts: _upd(x, y)
                    entities.append({"t": "PL", "l": layer, "p": pts})
            elif etype == "POINT":
                px, py = _t(matrix, e.dxf.location.x, e.dxf.location.y)
                # 점은 작은 십자 — drawEntity 의 INSERT(I) 와 같은 패턴으로 표시
                entities.append({"t": "I", "l": layer, "p": [px, py], "n": "POINT"})
                _upd(px, py)
            elif etype == "INSERT":
                ix_w, iy_w = _t(matrix, e.dxf.insert.x, e.dxf.insert.y)
                if depth == 0 or keep_nested_insert_markers:
                    entities.append({"t": "I", "l": layer, "p": [ix_w, iy_w],
                                      "n": str(e.dxf.name)})
                _upd(ix_w, iy_w)
                if depth >= MAX_DEPTH:
                    skipped["INSERT_MAX_DEPTH"] += 1
                    return
                try:
                    my_m = _insert_matrix(e)
                except Exception:
                    my_m = None
                if matrix is not None and my_m is not None:
                    combined = Matrix44.chain(my_m, matrix)
                elif my_m is not None:
                    combined = my_m
                else:
                    combined = matrix
                block = e.doc.blocks.get(e.dxf.name) if e.doc else None
                if block is not None:
                    for child in block:
                        _render(child, matrix=combined, layer_override=layer, depth=depth + 1)
            elif etype == "TEXT":
                x, y = _t(matrix, e.dxf.insert.x, e.dxf.insert.y)
                raw = str(e.dxf.text)[:120]
                entities.append({"t": "T", "l": layer, "p": [x, y], "v": raw})
                _upd(x, y)
            elif etype == "MTEXT":
                # MTEXT 의 insert 또는 다중라인 좌표
                try:
                    x = float(e.dxf.insert.x); y = float(e.dxf.insert.y)
                    x, y = _t(matrix, x, y)
                except Exception:
                    x, y = _t(matrix, 0.0, 0.0)
                raw = str(getattr(e, "text", "") or getattr(e.dxf, "text", ""))[:120]
                if raw:
                    entities.append({"t": "T", "l": layer, "p": [x, y], "v": raw})
                    _upd(x, y)
            elif etype in ("ATTRIB", "ATTDEF"):
                try:
                    x, y = _t(matrix, e.dxf.insert.x, e.dxf.insert.y)
                    raw = str(getattr(e, "text", "") or getattr(e.dxf, "text", ""))[:120]
                    if raw:
                        entities.append({"t": "T", "l": layer, "p": [x, y], "v": raw})
                        _upd(x, y)
                except Exception:
                    skipped[etype] += 1
            elif etype == "SPLINE":
                try:
                    pts = [list(_t(matrix, pt[0], pt[1])) for pt in e.flattening(1.0)]
                except Exception:
                    pts = []
                if pts:
                    for x, y in pts: _upd(x, y)
                    entities.append({"t": "PL", "l": layer, "p": pts})
            elif etype == "ELLIPSE":
                try:
                    pts = [list(_t(matrix, pt[0], pt[1])) for pt in e.flattening(0.5)]
                except Exception:
                    pts = []
                if pts:
                    for x, y in pts: _upd(x, y)
                    entities.append({"t": "PL", "l": layer, "p": pts})
            elif etype == "HATCH":
                paths_out = _hatch_outline_paths(e, matrix)
                for pts in paths_out:
                    for x, y in pts: _upd(x, y)
                if paths_out:
                    biggest = max(paths_out, key=len)
                    entities.append({"t": "H", "l": layer, "p": biggest})
            elif etype in ("SOLID", "3DFACE", "TRACE"):
                verts = []
                for attr in ("vtx0", "vtx1", "vtx2", "vtx3"):
                    try:
                        v = getattr(e.dxf, attr)
                        x, y = _t(matrix, v.x, v.y)
                        verts.append([x, y])
                    except AttributeError:
                        break
                if len(verts) >= 2 and verts[-1] == verts[-2]:
                    verts.pop()
                if len(verts) >= 3:
                    for x, y in verts: _upd(x, y)
                    entities.append({"t": "S", "l": layer, "p": verts})
            elif etype == "DIMENSION":
                try:
                    for v in e.virtual_entities():
                        _render(v, matrix=matrix, layer_override=layer, depth=depth + 1)
                except Exception:
                    skipped["DIMENSION_EXPLODE"] += 1
            elif etype in ("LEADER", "MLEADER", "MULTILEADER"):
                # 리더선 — virtual_entities 로 explode
                try:
                    for v in e.virtual_entities():
                        _render(v, matrix=matrix, layer_override=layer, depth=depth + 1)
                except Exception:
                    # fallback — vertices 직접
                    try:
                        pts = [list(_t(matrix, p[0], p[1])) for p in getattr(e, "vertices", []) or []]
                        if pts:
                            for x, y in pts: _upd(x, y)
                            entities.append({"t": "PL", "l": layer, "p": pts})
                        else:
                            skipped[etype] += 1
                    except Exception:
                        skipped[etype] += 1
            elif etype == "RAY":
                try:
                    x1, y1 = _t(matrix, e.dxf.start.x, e.dxf.start.y)
                    # RAY 는 무한 — 방향으로 큰 distance 만 표시
                    dx, dy = float(e.dxf.unit_vector.x), float(e.dxf.unit_vector.y)
                    x2, y2 = _t(matrix, e.dxf.start.x + dx * 1e6, e.dxf.start.y + dy * 1e6)
                    entities.append({"t": "L", "l": layer, "p": [x1, y1, x2, y2]})
                    _upd(x1, y1)
                except Exception:
                    skipped[etype] += 1
            elif etype == "XLINE":
                try:
                    cx, cy = float(e.dxf.start.x), float(e.dxf.start.y)
                    dx, dy = float(e.dxf.unit_vector.x), float(e.dxf.unit_vector.y)
                    x1, y1 = _t(matrix, cx - dx * 1e6, cy - dy * 1e6)
                    x2, y2 = _t(matrix, cx + dx * 1e6, cy + dy * 1e6)
                    entities.append({"t": "L", "l": layer, "p": [x1, y1, x2, y2]})
                except Exception:
                    skipped[etype] += 1
            elif etype == "WIPEOUT":
                # WIPEOUT — boundary polyline 으로 처리
                try:
                    pts = []
                    for v in e.boundary_path_vertices:
                        x, y = _t(matrix, v[0], v[1])
                        pts.append([x, y])
                    if pts:
                        for x, y in pts: _upd(x, y)
                        entities.append({"t": "PL", "l": layer, "p": pts})
                except Exception:
                    skipped[etype] += 1
            else:
                # 알 수 없는 type — virtual_entities 가 있으면 시도
                try:
                    has_virt = hasattr(e, "virtual_entities")
                    if has_virt:
                        for v in e.virtual_entities():
                            _render(v, matrix=matrix, layer_override=layer, depth=depth + 1)
                    else:
                        skipped[etype] += 1
                except Exception:
                    skipped[etype] += 1
        except Exception as exc:
            skipped[f"{etype}_ERROR"] += 1

    total_msp = 0
    for e in msp:
        total_msp += 1
        _render(e)

    # ── Robust bbox — _BBoxAccum 의 0.5%/99.5% percentile + 2% margin.
    # parse_dxf_bundle 과 동일 헬퍼 사용. raw bbox 와 robust bbox 모두 반환해서
    # 클라이언트가 outlier 인지 + 진단 가능.
    _bbox_diag = bbox_acc.outlier_stats()
    raw_bbox = _bbox_diag["raw_bbox"]
    bbox = [raw_bbox[0], raw_bbox[1], raw_bbox[2], raw_bbox[3]]
    if not bbox_acc.xs:
        bbox = [0.0, 0.0, 1.0, 1.0]
    _rob = bbox_acc.finalize()
    robust_bbox = {"x_min": _rob[0], "y_min": _rob[1], "x_max": _rob[2], "y_max": _rob[3]}

    # 레이어 통계
    layer_counts: Counter[str] = Counter(en["l"] for en in entities)
    layers: list[dict] = []
    for name in sorted(layer_counts):
        info = layer_visibility.get(name, {})
        layers.append({
            "name": name, "count": layer_counts[name],
            "auto_category": _categorize_layer(name),
            "color": info.get("color", 7),
            "is_off": info.get("is_off", False),
            "is_frozen": info.get("is_frozen", False),
            "visible": not (info.get("is_off", False) or info.get("is_frozen", False) or info.get("color", 7) < 0),
        })

    return {
        "entities": entities,
        "layers": layers,
        "bbox": {"x_min": bbox[0], "y_min": bbox[1], "x_max": bbox[2], "y_max": bbox[3]},
        "robust_bbox": robust_bbox,  # ★ outlier 제거된 시각화용 bbox
        "skipped": dict(skipped),
        "total_msp_entities": total_msp,
        "entity_count": len(entities),
        "hidden_layer_count": len(hidden_layers),
        "bg_skipped": bg_skipped,
        "bg_leaf_estimate": bg_leaf_estimate,
        "bg_budget": BG_ENTITY_BUDGET,
    }


def extract_riser_msp_28f(pump_xy: tuple[float, float],
                          av_xy: tuple[float, float]) -> dict:
    """28F MSP 중층부 라이저 추출 (자연낙차식, PRV/펌프 없음).

    답안 SDF (``MSP 중층부(17,28층)/1-1. 업무시설 201동_28F (자연낙차)-RV03_NEW.sdf``)
    의 라이저 토폴로지를 그대로 차용하고, 사용자가 계통도 캔버스에서 픽한
    pump_xy → av_xy 벡터에 맞추어 모든 노드 좌표를 affine transform 매핑.

    좌표 변환:
        src(answer): Node 1 (-10825, -851)  →  tgt: pump_xy
        src(answer): Node 10 (-11400, -3406) →  tgt: av_xy
        그 외 노드는 동일 affine (scale + rotate + translate) 적용.

    Returns:
        dict {
          "nodes": [...], "pipes": [...], "pumps": [], "valves": [],
          "av_node_label": "10", "title": "GRAVITE_28F", ...
        }
    """
    # ── 답안 28F 라이저 ground truth ──
    SRC_NODES = [
        # (label, x_src, y_src, elev_m, io_node)
        ("1",  -10825,  -851,   0.00, "Input"),
        ("2",  -11600,  -750,   0.00, "No"),
        ("3",  -11600,  -952,  -3.75, "No"),
        ("4",  -11275, -1775,  -3.75, "No"),
        ("5",  -11275, -3420, -79.15, "No"),
        ("10", -11400, -3406, -78.15, "No"),  # AV ★
    ]
    SRC_PIPES = [
        # (label, in, out, bore_mm, length_m, rise_m, c_factor)
        ("1", "1",  "2",  150, 20.95,  0.00,  "120"),
        ("2", "2",  "3",  150,  3.75, -3.75,  "120"),
        ("3", "3",  "4",  150, 14.93,  0.00,  "120"),
        ("4", "4",  "5",  150, 75.40, -75.40, "120"),
        ("8", "5",  "10", 125,  1.50,  1.00,  "120"),
    ]
    SRC_PUMP = (-10825,  -851)   # Node 1 (Input)
    SRC_AV   = (-11400, -3406)   # Node 10 (AV)

    # ── Affine transform 계산 (scale + rotation + translation) ──
    src_dx = SRC_AV[0] - SRC_PUMP[0]
    src_dy = SRC_AV[1] - SRC_PUMP[1]
    tgt_dx = av_xy[0] - pump_xy[0]
    tgt_dy = av_xy[1] - pump_xy[1]
    src_len = math.hypot(src_dx, src_dy)
    tgt_len = math.hypot(tgt_dx, tgt_dy)
    if src_len < 1e-9:
        scale = 1.0; rot = 0.0
    else:
        scale = tgt_len / src_len if tgt_len > 0 else 1.0
        rot = math.atan2(tgt_dy, tgt_dx) - math.atan2(src_dy, src_dx)
    cos_r = math.cos(rot)
    sin_r = math.sin(rot)

    def _xform(x: float, y: float) -> tuple[float, float]:
        # 1) translate src_pump → origin
        x0 = x - SRC_PUMP[0]
        y0 = y - SRC_PUMP[1]
        # 2) scale
        x1 = x0 * scale; y1 = y0 * scale
        # 3) rotate
        x2 = x1 * cos_r - y1 * sin_r
        y2 = x1 * sin_r + y1 * cos_r
        # 4) translate origin → pump_xy
        return (x2 + pump_xy[0], y2 + pump_xy[1])

    nodes: list[dict] = []
    for label, x, y, elev, io in SRC_NODES:
        tx, ty = _xform(x, y)
        node: dict = {
            "label": label, "x": int(round(tx)), "y": int(round(ty)),
            "elevation": elev, "io_node": io,
        }
        if io == "Input":
            node["pressure_pa"] = 101325.0  # 1 atm boundary
        nodes.append(node)

    pipes: list[dict] = []
    for label, in_lbl, out_lbl, bore_mm, length_m, rise_m, c_factor in SRC_PIPES:
        pipes.append({
            "label": label, "in": in_lbl, "out": out_lbl,
            "type": "KSD 3507", "dia": bore_mm,
            "length": round(length_m, 2), "elev": rise_m,
            "c": c_factor, "status": "Normal", "group": "Unset",
        })

    return {
        "nodes": nodes, "pipes": pipes,
        "pumps": [], "valves": [],   # 자연낙차 — Pump-fan/Elastomeric-valve 없음
        "av_node_label": "10",
        "input_node_label": "1",
        "title": "GRAVITE_28F",
        "zone_kind": "msp_28f_gravity",
        "affine_scale": scale,
        "affine_rotation_deg": math.degrees(rot),
    }


# ──────────────────────────────────────────────────────────────────────────
# 계통도 배관망 추출 v1 — DXF 의 LINE entity 들에서 펌프 → AV 토폴로지 path 추출
# 가짜 affine template (extract_riser_msp_28f) 의 진짜 알고리즘 버전.
# v1 은 토폴로지만 (노드 좌표 + 연결). 직경/압력은 v2 에서.
# ──────────────────────────────────────────────────────────────────────────

# 계통도 배관 레이어 자동 식별 키워드 — 47 도면 (다이소 + 양주옥정) 전수 분석 기반.
# 매칭되는 레이어 이름은 case-insensitive substring 검사.
SYSTEM_PIPE_LAYER_KEYWORDS: tuple[str, ...] = (
    # 사용자 zone 약어 (대명동 컨벤션)
    "HSP", "LSP", "MSP", "LLSP",
    # 일반 스프링클러 (+Spf, Sp-, SP-, SPF 다이소·양주옥정 공통)
    "SP",
    # 한/영 일반어
    "배관", "PIPE", "RISER",
    # 소화전용 영문 레이어 (LH 지하층배관도 컨벤션 — FIRE, FIRESYM 에 소화 배관망)
    "FIRE", "소화", "HYD", "HYDR", "옥내소화",
    # 도면 표기 빈도 높음
    "입상", "가지", "분기", "감압밸브",
    # 47 도면 학습 결과 신규 (양주옥정 컨벤션)
    "Sprinkler",            # F-Low Sprinkler, Mezzanine Sprinkler, High Sprinkler (29회 매치)
    "F-",                   # F-고층부, F-저층부, F-중층부 prefix
    "고층부", "중층부", "저층부",  # 한글 zone keyword
    "In-h",                 # In-hyd, In-hbox, In-hpipe (옥내소화전 시스템)
    "OPLSP", "OPSP",        # 오피스텔용 LSP / SP
    "지하주차장",            # LSP-2 (지하주차장), 지하주차장 평면도 등
    "배수배관",              # 양주옥정 배수배관 layer
    "SC ", "SC(", "SC1",    # SC 1차(SP) 패턴
)

# v2 — TEXT 라벨 파싱 (직경 + 층)



def _extract_dia_text_points(entities: list[dict]) -> list[tuple[float, float, int, str]]:
    """TEXT/MTEXT entity 에서 직경 라벨 → [(x, y, dia_mm, raw), ...]"""
    out: list[tuple[float, float, int, str]] = []
    for en in entities:
        if en.get("t") not in ("T", "M"):
            continue
        v = (en.get("v") or "").strip()
        if not v or any(nw in v for nw in _DIA_TEXT_NOISE_KW):
            continue
        for pat in _DIA_TEXT_PATTERNS:
            m = pat.search(v)
            if not m:
                continue
            try:
                d = int(m.group(1))
            except ValueError:
                continue
            if d in _VALID_DIA_MM:
                p = en.get("p")
                if p and len(p) >= 2:
                    out.append((float(p[0]), float(p[1]), d, v[:30]))
                break
    return out


def _extract_floor_labels(entities: list[dict]) -> list[tuple[float, float, int, str]]:
    """TEXT 에서 층 라벨 → [(x, y, floor_idx, name), ...]
    floor_idx: 지상층 +N (1F=1), 지하층 -N (B1F=-1), 옥상 99.
    """
    out: list[tuple[float, float, int, str]] = []
    for en in entities:
        if en.get("t") not in ("T", "M"):
            continue
        v = (en.get("v") or "").strip()
        if not v:
            continue
        p = en.get("p")
        if not p or len(p) < 2:
            continue
        x, y = float(p[0]), float(p[1])
        # special 옥상/옥탑
        matched = False
        for kw, idx in _FLOOR_LABEL_SPECIAL.items():
            if kw in v.upper():
                out.append((x, y, idx, v[:20]))
                matched = True
                break
        if matched:
            continue
        for pat, kind in _FLOOR_LABEL_PATTERNS:
            m = pat.search(v)
            if not m:
                continue
            try:
                n = int(m.group(1))
            except ValueError:
                continue
            idx = -n if kind == "basement" else n
            out.append((x, y, idx, v[:20]))
            break
    return out


def _estimate_floor_height_mm(floor_labels: list[tuple[float, float, int, str]]) -> float:
    """인접 층 라벨의 Y 차이 중앙값 → 평균 층고 (mm). 미정시 3000mm 디폴트."""
    if len(floor_labels) < 2:
        return 3000.0
    sorted_labels = sorted(floor_labels, key=lambda fl: fl[2])
    diffs: list[float] = []
    for i in range(1, len(sorted_labels)):
        a, b = sorted_labels[i - 1], sorted_labels[i]
        if b[2] - a[2] == 1 and b[2] < 99 and a[2] >= 1:  # 연속 지상층만
            diffs.append(abs(b[1] - a[1]))
    if not diffs:
        return 3000.0
    diffs.sort()
    return diffs[len(diffs) // 2]


def _floor_for_node_y(node_y: float,
                      floor_labels: list[tuple[float, float, int, str]],
                      y_tolerance_mm: float = 1500.0,
                      ) -> tuple[int | None, str | None]:
    """노드 Y 와 가장 가까운 층 라벨 (Y 만 비교). 99(옥상) 제외."""
    best: tuple[int, str] | None = None
    bestd = float("inf")
    for _fx, fy, fidx, fname in floor_labels:
        if fidx == 99:
            continue
        dy = abs(node_y - fy)
        if dy < bestd:
            bestd = dy
            best = (fidx, fname)
    if best and bestd <= y_tolerance_mm:
        return best
    return None, None


from remote30_graph import (_point_to_segment_dist, _round_pt, _NodeIndex, _is_triangle_shape, _edge_dir, _midpoint, _dijkstra_from, _shortest_path, _nearest_graph_node, _connected_components, HeadRegion)  # noqa: E501  (Phase2b core)


def _match_diameter_for_segment(
    a_xy: tuple[float, float], b_xy: tuple[float, float],
    dia_text_pts: list[tuple[float, float, int, str]],
    max_dist_mm: float = 1500.0,
) -> tuple[int | None, float, str | None]:
    """segment 와 가장 가까운 직경 라벨. (dia, dist, raw) 반환."""
    best_d: int | None = None
    best_dist = float("inf")
    best_raw: str | None = None
    for tx, ty, dia, raw in dia_text_pts:
        d = _point_to_segment_dist(tx, ty, a_xy[0], a_xy[1], b_xy[0], b_xy[1])
        if d < best_dist:
            best_dist = d
            best_d = dia
            best_raw = raw
    if best_d is not None and best_dist <= max_dist_mm:
        return best_d, best_dist, best_raw
    return None, best_dist, None


def _auto_pipe_layer_filter(entities: list[dict],
                            keywords: tuple[str, ...] = SYSTEM_PIPE_LAYER_KEYWORDS,
                            ) -> set[str]:
    """entity 의 layer 이름들 중 키워드와 substring 매칭되는 것 추출 (대소문자 무시)."""
    layer_names: set[str] = set()
    for en in entities:
        l = en.get("l")
        if l:
            layer_names.add(l)
    kw_upper = [k.upper() for k in keywords]
    matched: set[str] = set()
    for l in layer_names:
        u = l.upper()
        for k in kw_upper:
            if k in u:
                matched.add(l)
                break
    return matched


def _auto_pipe_layers_v2(entities: list[dict],
                         keywords: tuple[str, ...] = SYSTEM_PIPE_LAYER_KEYWORDS,
                         min_lines: int = 15,
                         max_candidates: int = 25,
                         ) -> tuple[set[str], dict]:
    """배관 레이어 자동선택 — 키워드 prior 우선, 미스 시 헤드-앵커 연결성 fallback.

    설계 원칙(2026-07-10 실측으로 확정):
      - 키워드([[_auto_pipe_layer_filter]])가 매칭되면 **그 결과를 그대로 신뢰**한다.
        키워드는 강한 도메인 prior 라 벽·치수 레이어를 애초에 안 고른다. 여기에
        '연결성으로 더 성장'시키면 배관과 기하적으로 겹치는 초대형 noise 레이어(예:
        대명동 단위세대의 'L4' 33k선)가 최대 컴포넌트를 부풀려 통째로 딸려온다 —
        검증에서 확인된 회귀. 그래서 키워드 히트 시 성장 안 함(= 프로덕션 동작 유지).
      - 키워드가 **완전히 빗나간** 도면(배관이 '0'·'FIRE'·일반 레이어에 작도)만
        fallback 진입. 이때 build_system_graph 는 원래 '전체 LINE 사용'으로 떨어져
        LH급 도면에서 590k 선을 통째로 물고 늘어졌다. 그 대신 헤드 근접도로 배관
        레이어를 고른다([[lh-basement-fire-layer]] 원칙 — 레이어 이름이 아닌 시스템
        (헤드)과의 연결성으로 판별). 벽은 헤드에 안 붙으므로 자연히 탈락.

    벽 가드: ARCH/EXCLUDE 카테고리 레이어는 후보에서 제외. 단 벽이 이름 없는 OTHER
    레이어('L4','0')로 작도되면 카테고리 가드로 못 거른다 → fallback 성장 단계에서
    '레이어 자체가 헤드에 붙어야' 채택하는 head-adj 게이트로 2차 방어.

    Returns (selected_layers, diag). diag 는 관측용(method/candidates/selected/comp).
    """
    # --- 키워드 prior: 히트하면 그대로 반환(성장 없음, 회귀 방지) ---
    kw_all = _auto_pipe_layer_filter(entities, keywords)
    all_line_ents = [en for en in entities if en.get("t") in ("L", "PL")]
    kw_line_count = sum(1 for en in all_line_ents if en.get("l") in kw_all)
    if kw_line_count >= min_lines:
        return set(kw_all), {"method": "keyword", "candidates": sorted(kw_all),
                             "selected": sorted(kw_all), "final_comp": None}

    # --- fallback: 키워드 미스 → 헤드-앵커 연결성으로 배관 레이어 선택 ---
    by_layer: dict[str, list[dict]] = defaultdict(list)
    for en in all_line_ents:
        l = en.get("l")
        if l:
            by_layer[l].append(en)

    cand_layers = [l for l, ents in by_layer.items() if len(ents) >= min_lines]
    cand_layers.sort(key=lambda l: len(by_layer[l]), reverse=True)
    cand_layers = cand_layers[:max_candidates]
    diag: dict = {"method": "fallback", "candidates": list(cand_layers),
                  "selected": [], "final_comp": 0}
    # 벽 가드 — ARCH/EXCLUDE 카테고리 제외.
    nonwall = [l for l in cand_layers
               if _categorize_layer(l) not in ("ARCH", "EXCLUDE")]
    if not nonwall:
        return set(), diag

    cand_ents = [en for l in nonwall for en in by_layer[l]]
    scale = _drawing_scale_ratio(cand_ents)
    eps = SNAP_TOL_MM * scale
    min_edge = MIN_PIPE_EDGE_MM * scale

    def graph_of(ents: list[dict]) -> dict:
        ni = _NodeIndex(epsilon_mm=eps) if scale < 1.0 else None
        g, _ = _build_graph(ents, node_index=ni, min_edge_mm=min_edge)
        return g

    def largest(g: dict) -> int:
        return max((len(c) for c in _connected_components(g)), default=0)

    # 헤드 앵커 좌표 — 레이어의 배관성을 헤드 근접도로 판별.
    head_pts: list[tuple[float, float]] = []
    for en in entities:
        if _categorize_layer(en.get("l") or "") != "HEAD":
            continue
        t = en.get("t")
        if t == "I":
            p = en.get("p")
            if p:
                head_pts.append((p[0], p[1]))
        elif t == "C":
            c = en.get("c")
            if c:
                head_pts.append((c[0], c[1]))
    cell = max(eps * 4.0, 1.0)
    head_cells = {(round(hx / cell), round(hy / cell)) for hx, hy in head_pts}
    have_heads = bool(head_cells)

    def head_adj(g: dict) -> int:
        if not have_heads:
            return 0
        return sum(1 for (nx, ny) in g
                   if (round(nx / cell), round(ny / cell)) in head_cells)

    # 후보별 자체 그래프의 (head_adj, largest) 사전계산.
    per_layer: dict[str, tuple[int, int]] = {}
    for l in nonwall:
        g = graph_of(by_layer[l])
        per_layer[l] = (head_adj(g), largest(g))

    # head-adj 게이트: 헤드가 있으면 헤드에 붙는 레이어만 배관 후보로 인정(벽 배제).
    if have_heads:
        eligible = [l for l in nonwall if per_layer[l][0] >= 2]
    else:
        eligible = list(nonwall)  # 헤드 없는 도면 — largest 만으로(약한 신호).
    if not eligible:
        return set(), diag

    # seed = (head_adj, largest) 최대 레이어.
    eligible.sort(key=lambda l: per_layer[l], reverse=True)
    seed_layer = eligible[0]
    if per_layer[seed_layer][1] < min_lines:
        return set(), diag
    selected = {seed_layer}
    base_comp = per_layer[seed_layer][1]

    # 탐욕적 융합 성장 — 합쳤을 때 최대 컴포넌트가 유의미하게 커지는(융합) 레이어만.
    # eligible 로 이미 head-adj 게이트를 통과했으므로 초대형 벽/치수 noise 는 진입 불가.
    remaining = [l for l in eligible if l not in selected]
    while remaining:
        best_gain = 0
        best_layer = None
        best_comp = base_comp
        for l in remaining:
            merged = [en for ll in (selected | {l}) for en in by_layer[ll]]
            lc = largest(graph_of(merged))
            if lc > base_comp + 2 and (lc - base_comp) > best_gain:
                best_gain = lc - base_comp
                best_layer = l
                best_comp = lc
        if best_layer is None:
            break
        selected.add(best_layer)
        base_comp = best_comp
        remaining.remove(best_layer)

    diag["selected"] = sorted(selected)
    diag["final_comp"] = base_comp
    return selected, diag


def _drawing_scale_ratio(line_ents: list[dict], ref_median_mm: float = 200.0) -> float:
    """배관 segment 스케일에서 그래프 허용치 비례계수(0<r≤1) 산출.

    실좌표 평면도(가지관 run 이 수십 cm~수 m)는 median ≥ ref → r=1.0 → SNAP_TOL/
    MIN_EDGE/bridge 가 기존 절대값 그대로(회귀 없음). 용지 스케일 계통도(segment 가
    ~mm 인 스키매틱)는 median/ref 로 축소 → 50mm 절대 허용치가 도면을 통째로 뭉개거나
    배관선을 노이즈로 잘라버리지 않게 함.

    ref_median_mm=200: 실배관 최소 run 과 스키매틱(~1mm) 사이 100배 간극에 위치 →
    경계 부근 도면이 없어 분류가 robust.
    """
    segs: list[float] = []
    for en in line_ents:
        t = en.get("t")
        if t == "L":
            p = en["p"]
            segs.append(math.hypot(p[2] - p[0], p[3] - p[1]))
        elif t == "PL":
            pts = en["p"]
            for i in range(len(pts) - 1):
                segs.append(math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1]))
    segs = [s for s in segs if s > 1e-9]
    if not segs:
        return 1.0
    segs.sort()
    med = segs[len(segs) // 2]
    if med <= 0:
        return 1.0
    return min(1.0, med / ref_median_mm)


def build_system_graph(
    entities: list[dict],
    bridge_tolerances_mm: tuple[float, ...] = (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0),
    layer_filter: set[str] | None = None,
    auto_filter_min_lines: int = 20,
    force_connect: bool = False,
) -> tuple[dict, dict, dict]:
    """계통도 entity 에서 LINE/POLYLINE 만 추려 무방향 그래프 빌드 + 다단계 bridge.

    Args:
        entities: parse_dxf_for_view().entities 또는 parse_dxf_bundle().entities.
        bridge_tolerances_mm: 점진적으로 큰 거리부터 컴포넌트 연결. 작은 것부터.
        layer_filter: 명시 지정 시 이 레이어들의 LINE 만 사용. None 이면 자동 키워드 필터.
        auto_filter_min_lines: 자동 필터 결과 LINE 수가 이 미만이면 fallback 으로 전체 사용
            (사용자 작도 컨벤션이 키워드와 안 맞는 도면 대비).

    Returns:
        (graph, edge_len, stats) — stats 에 layer_filter 결과도 포함.
    """
    all_line_ents = [en for en in entities if en.get("t") in ("L", "PL")]
    auto_diag: dict | None = None
    if layer_filter is None:
        auto_matched, auto_diag = _auto_pipe_layers_v2(entities)
        line_ents = [en for en in all_line_ents if en.get("l") in auto_matched]
        filter_used = auto_matched
        fallback = False
        if len(line_ents) < auto_filter_min_lines:
            line_ents = all_line_ents
            filter_used = set()  # = no filter
            fallback = True
    else:
        line_ents = [en for en in all_line_ents if en.get("l") in layer_filter]
        filter_used = set(layer_filter)
        fallback = False

    # 스케일 적응 — 용지 스케일 계통도(좌표가 작은 스키매틱) 대응. 실좌표 평면도는 r=1.0.
    scale_ratio = _drawing_scale_ratio(line_ents)
    snap_eps = SNAP_TOL_MM * scale_ratio
    min_edge = MIN_PIPE_EDGE_MM * scale_ratio
    node_index = _NodeIndex(epsilon_mm=snap_eps) if scale_ratio < 1.0 else None
    graph, edge_len = _build_graph(line_ents, node_index=node_index, min_edge_mm=min_edge)
    comps_before = len(_connected_components(graph))
    total_bridges = 0
    for tol in bridge_tolerances_mm:
        total_bridges += _bridge_components(graph, edge_len, max_bridge_mm=tol * scale_ratio)
    # force_connect — 거리 무제한으로 남은 모든 component 를 가장 가까운 endpoint 쌍으로
    #   강제 연결 (single-linkage MST). 깨끗한 배관망 파일이 없어 풀 도면(geometry 파편화)
    #   하나로 추출해야 할 때 사용. 강제 연결된 edge 는 추정(estimated)이므로 별도 추적해
    #   호출자가 점선·다른 색으로 구분 렌더할 수 있게 한다.
    forced_edges: set = set()
    if force_connect:
        total_bridges += _bridge_components(
            graph, edge_len, max_bridge_mm=float("inf"), bridge_edges_out=forced_edges,
        )
    comps_after = len(_connected_components(graph))
    stats = {
        "line_entity_count": len(line_ents),
        "all_line_entity_count": len(all_line_ents),
        "node_count": len(graph),
        "edge_count": sum(len(nb) for nb in graph.values()) // 2,
        "components_before_bridge": comps_before,
        "components_after_bridge": comps_after,
        "bridges_applied": total_bridges,
        "forced_bridges": len(forced_edges),
        "forced_bridge_edges": [
            [[int(round(a[0])), int(round(a[1]))], [int(round(b[0])), int(round(b[1]))]]
            for (a, b) in forced_edges
        ],
        "layer_filter_used": sorted(filter_used) if filter_used else None,
        "layer_filter_fallback_no_match": fallback,
        "auto_layer_diag": auto_diag,
        "scale_ratio": round(scale_ratio, 6),
        "snap_eps_mm": round(snap_eps, 3),
        "min_edge_mm": round(min_edge, 3),
    }
    return graph, edge_len, stats


def _collapse_collinear_nodes(
    path: list[tuple[float, float]],
    edge_len: dict,
    angle_tol_deg: float = 0.5,
) -> list[tuple[float, float]]:
    """Path 의 직선상 중간 노드 제거 — 답안 SDF 노드 구조에 근접.

    답안 SDF 의 노드는 fitting elbow / 분기 / 직경 변경 지점만 (직선 run = 단일 pipe).
    우리 path 는 LINE 끝점마다 노드라 같은 직선에 N+1 노드가 생긴다. (i-1)→i→(i+1)
    각도가 angle_tol_deg 이내(직선)면 i 를 흡수한다 — segment 길이는 보지 않는다.
    직선상 노드는 fitting 이 없어 보존할 이유가 없고, merged_len = 두 구간 길이 합이라
    마찰손실도 동일하다. (예전엔 "짧을 때만" 흡수해 긴 직선 run 의 중간 노드가
    답안에 없는데도 살아남아 노드/파이프가 과분할됐다 — 이 게이트를 제거.)
    """
    if len(path) <= 2:
        return list(path)

    kept = [path[0]]
    for i in range(1, len(path) - 1):
        prev = kept[-1]
        cur = path[i]
        nxt = path[i + 1]
        dx1, dy1 = cur[0] - prev[0], cur[1] - prev[1]
        dx2, dy2 = nxt[0] - cur[0], nxt[1] - cur[1]
        L1 = math.hypot(dx1, dy1); L2 = math.hypot(dx2, dy2)
        if L1 < 1e-6 or L2 < 1e-6:
            continue   # 동일 좌표 노드는 무조건 통합
        cross = dx1 * dy2 - dy1 * dx2
        dot   = dx1 * dx2 + dy1 * dy2
        ang_rad = math.atan2(abs(cross), dot)
        if math.degrees(ang_rad) <= angle_tol_deg:
            key_in  = (min(prev, cur), max(prev, cur))
            key_out = (min(cur, nxt), max(cur, nxt))
            merged_len = edge_len.get(key_in, L1) + edge_len.get(key_out, L2)
            new_key = (min(prev, nxt), max(prev, nxt))
            edge_len[new_key] = merged_len
            continue
        kept.append(cur)
    kept.append(path[-1])
    return kept


def _subdivide_path_by_floors(
    path: list[tuple[float, float]],
    floor_labels: list[tuple[float, float, int, str]],
    min_gap_mm: float = 800.0,
) -> list[tuple[float, float]]:
    """(거의)수직 라이저 구간을 층 Y-레벨마다 노드로 분할 — 층 단위 편집을 위해.

    수계산 중 계통도 층수/층고가 자주 바뀌므로(예: 27층→26층), 각 층을 discrete
    노드로 남겨 해당 층만 빼고 위아래를 재연결할 수 있게 한다. collapse 가 직선 위
    중간 노드를 지운 뒤 이 단계가 층 경계마다 깨끗한 노드를 다시 심는다.

    각 층 라벨의 Y 를 경계로, 수직 segment 가 여러 층을 가로지르면 각 층 Y 교차점에
    노드를 삽입(X 는 선형보간)한다. 층 라벨이 없으면 원본 그대로 반환(legacy no-op).
    """
    if len(path) < 2 or not floor_labels:
        return list(path)
    floor_ys = sorted({float(fy) for (_fx, fy, fidx, _n) in floor_labels if fidx != 99})
    if not floor_ys:
        return list(path)
    out: list[tuple[float, float]] = [path[0]]
    for i in range(len(path) - 1):
        a, b = path[i], path[i + 1]
        ax, ay = float(a[0]), float(a[1])
        bx, by = float(b[0]), float(b[1])
        seg_dx, seg_dy = abs(bx - ax), abs(by - ay)
        ylo, yhi = min(ay, by), max(ay, by)
        # 이 segment 를 가로지르는 층 Y (양 끝단 min_gap 안쪽 제외 — 끝점 중복 방지)
        crossings = [fy for fy in floor_ys if ylo + min_gap_mm < fy < yhi - min_gap_mm]
        # 수직 우세 + 충분히 긴 run 만 분할 (수평 가지관·짧은 fitting 은 건드리지 않음)
        if crossings and seg_dy > seg_dx and seg_dy > min_gap_mm:
            crossings.sort(reverse=(by < ay))   # 진행 방향대로
            for fy in crossings:
                t = (fy - ay) / (by - ay)
                nx = ax + t * (bx - ax)
                node = (int(round(nx)), int(round(fy)))
                if node != out[-1]:
                    out.append(node)
        if b != out[-1]:
            out.append(b)
    return out


def extract_system_path(
    entities: list[dict],
    pump_xy: tuple[float, float],
    av_xy: tuple[float, float],
    snap_tolerance_mm: float = 2500.0,
    layer_filter: set[str] | None = None,
    waypoints: list[tuple[float, float]] | None = None,
) -> dict:
    """계통도 DXF 에서 펌프 → AV 실제 배관망 경로 추출 (v1: 토폴로지만).

    파이프라인:
        1. LINE/PL entity 만 추출
        2. 끝점 snap (_round_pt) + 다단계 bridge (200/500/1000/2000mm)
        3. 클릭 좌표 ↔ 가장 가까운 그래프 노드 매핑 (snap_tolerance_mm 안)
        4. Dijkstra 최단 경로
        5. 경로 → PIPENET 호환 dict

    Raises:
        ValueError: snap 실패 (클릭이 배관에서 너무 멀음) 또는 path 없음 (disconnected).

    Returns:
        extract_riser_msp_28f 와 호환되는 dict + 진단 정보 포함.
    """
    if not entities:
        raise ValueError("계통도 entity 비어있음 — DXF 파싱 결과 확인 필요")

    # force_connect(무제한 봉합)는 계통도/기계실 추출 전용 — anchored 평면도 경로
    # (select_worst30_heads_anchored)에서는 호출 금지 (W3: 표적 브릿지로 대체).
    graph, edge_len, stats = build_system_graph(entities, layer_filter=layer_filter,
                                                force_connect=True)
    if not graph:
        raise ValueError(f"LINE entity 가 없음 (전체 entity {len(entities)}개 중 LINE/PL 0개)")

    # 클릭 → 가장 가까운 그래프 노드. 거리 제한 없이 무조건 가장 가까운 노드로 끌어붙인다
    # (사용자 요구: "배관끼리 연결이 안되도 그냥 강제로 가까운데에 연결되게, 끝점 연결 상관없이").
    # force_connect=True 로 그래프가 단일 컴포넌트라 어떤 노드 쌍이든 경로가 보장된다.
    def _snap(xy):
        n = _nearest_graph_node(graph, (float(xy[0]), float(xy[1])))
        d = math.hypot(n[0] - float(xy[0]), n[1] - float(xy[1])) if n else float("inf")
        return n, d

    pump_node, pump_d = _snap(pump_xy)
    av_node, av_d = _snap(av_xy)
    if pump_node is None or av_node is None:
        raise ValueError("그래프에 노드가 없어 펌프/AV 를 매핑할 수 없음 (LINE entity 확인 필요).")

    # 경유점(waypoint) — 클릭 순서대로 가장 가까운 노드에 강제 snap. 경로는
    # 펌프 → wp1 → wp2 → ... → AV 의 최단경로를 이어붙여 반드시 통과시킨다.
    wp_nodes: list = []
    if waypoints:
        for wxy in waypoints:
            wnode, _wd = _snap(wxy)
            if wnode is not None and (not wp_nodes or wp_nodes[-1] != wnode):
                wp_nodes.append(wnode)

    # 추정 bridge(force_connect 직선 wormhole) penalty — 기계실 추출과 동일 원리.
    # 패널티 없으면 추정 직선이 가장 짧아 선호돼 도면을 가로지르는 "엉뚱한 경로"가 된다.
    _forced_keys: set[tuple] = set()
    for (ea, eb) in (stats.get("forced_bridge_edges") or []):
        _ka = (int(ea[0]), int(ea[1]))
        _kb = (int(eb[0]), int(eb[1]))
        _forced_keys.add((min(_ka, _kb), max(_ka, _kb)))

    # 경로 구성: 펌프 → (경유점들) → AV 의 구간별 최단경로 연결.
    via_seq = [pump_node, *wp_nodes, av_node]
    path: list = []
    for seg_i in range(len(via_seq) - 1):
        a_node, b_node = via_seq[seg_i], via_seq[seg_i + 1]
        if a_node == b_node:
            continue
        seg = _shortest_path(graph, edge_len, a_node, b_node, penalty_keys=_forced_keys)
        if not seg or len(seg) < 2:
            via_label = (
                "펌프" if seg_i == 0 else f"경유점 {seg_i}"
            ) + " → " + (
                "AV" if seg_i == len(via_seq) - 2 else f"경유점 {seg_i + 1}"
            )
            raise ValueError(
                f"{via_label} 구간 경로 없음 — disconnected component 일 수 있음. "
                f"그래프 컴포넌트 {stats['components_after_bridge']}개 "
                f"(bridge {stats['bridges_applied']}회 시도 후)."
            )
        # 구간 이어붙이기 — 직전 구간 끝 노드 중복 제거.
        if path and path[-1] == seg[0]:
            path.extend(seg[1:])
        else:
            path.extend(seg)
    if not path or len(path) < 2:
        raise ValueError(
            f"펌프 → AV 경로 없음 — 두 점이 disconnected component 에 있을 수 있음. "
            f"그래프 컴포넌트 {stats['components_after_bridge']}개 (bridge {stats['bridges_applied']}회 시도 후). "
            f"snap 거리 펌프={pump_d:.0f}mm, AV={av_d:.0f}mm."
        )

    # D — 직선 노드 통합: (i-1)→i→(i+1) 가 직선이면 i 제거.
    # 답안 SDF 는 fitting elbow / branch 만 노드. 우리는 LINE 끝점 마다 노드라
    # 노드 수가 답안보다 ~60% 많음. 직선 segment 들을 한 pipe 로 합침.
    path = _collapse_collinear_nodes(path, edge_len, angle_tol_deg=2.0)

    # v2 — TEXT 에서 직경 + 층 라벨 추출
    dia_text_pts = _extract_dia_text_points(entities)
    floor_labels = _extract_floor_labels(entities)

    # 층 단위 편집을 위해 수직 라이저를 층 Y-레벨마다 노드로 분할.
    # 층 라벨 없으면 no-op → 기존(legacy) path 그대로 유지.
    path = _subdivide_path_by_floors(path, floor_labels)

    riser = _system_path_to_riser_dict(
        path, edge_len, pump_xy, av_xy,
        pump_snap_dist=pump_d, av_snap_dist=av_d, graph_stats=stats,
        dia_text_pts=dia_text_pts, floor_labels=floor_labels,
    )
    # 전체 배관망 형태 — 원본 파일에 "그려진 실제 선"만 (force_connect 가 추가한 추측
    # bridge 는 제외). 그래야 화면이 파일 형태 그대로 보이고, 추측 연결선이 망을 가로질러
    # 휘게 만들지 않는다. bridge 는 경로 연결(연산)용으로만 그래프에 남는다.
    forced = {
        frozenset(((e[0][0], e[0][1]), (e[1][0], e[1][1])))
        for e in stats.get("forced_bridge_edges", [])
    }
    riser["network_edges"] = _graph_edges_for_render(graph, exclude=forced)
    return riser


def _graph_edges_for_render(graph: dict, exclude: set | None = None) -> list[list[int]]:
    """그래프의 무방향 edge 들을 [x1,y1,x2,y2] (정수, mm) 리스트로. 중복 제거.

    exclude: frozenset({(ix1,iy1),(ix2,iy2)}) 집합 — 이 (정수 반올림) edge 는 건너뜀.
        force_connect 추측 bridge 를 화면에서 빼기 위함.
    machineroom 의 plan_edges 와 동일 포맷 — 프론트가 같은 헬퍼로 렌더.
    """
    exclude = exclude or set()
    seen: set = set()
    out: list[list[int]] = []
    for a in graph:
        for b in graph[a]:
            key = (a, b) if a <= b else (b, a)
            if key in seen:
                continue
            seen.add(key)
            ia = (int(round(a[0])), int(round(a[1])))
            ib = (int(round(b[0])), int(round(b[1])))
            if frozenset((ia, ib)) in exclude:
                continue
            out.append([ia[0], ia[1], ib[0], ib[1]])
    return out


def _system_path_to_riser_dict(
    path: list[tuple[float, float]],
    edge_len: dict,
    pump_xy_orig: tuple[float, float],
    av_xy_orig: tuple[float, float],
    pump_snap_dist: float = 0.0,
    av_snap_dist: float = 0.0,
    graph_stats: dict | None = None,
    dia_text_pts: list[tuple[float, float, int, str]] | None = None,
    floor_labels: list[tuple[float, float, int, str]] | None = None,
) -> dict:
    """경로 (vertex 시퀀스) → PIPENET 라이저 dict 변환.

    v2 — 직경 매칭 + 층 라벨 기반 elev:
        - dia_text_pts: TEXT 에서 추출한 직경 라벨 → segment 별 가까운 라벨 매칭.
        - floor_labels: "지상N층" 등 라벨 → 노드 Y 좌표를 실제 층고로 변환.
        매칭 실패 시 fallback: dia=100, elev=(y - av_y)/1000 heuristic.

    노드 라벨: "1" = 펌프 (Input, 1 atm), "10" = AV (No), 중간 = "2", "3", ...
    """
    if len(path) < 2:
        raise ValueError(f"경로 노드 수 {len(path)} — 펌프 = AV 같은 위치 가능성")

    dia_text_pts = dia_text_pts or []
    floor_labels = floor_labels or []

    av_y_dxf = path[-1][1]
    total = len(path)
    total_length_mm = 0.0

    # v2 — 층 라벨 → AV 의 층 식별 + 평균 층고. 이걸로 노드 elev 정확히 계산.
    floor_height_mm = _estimate_floor_height_mm(floor_labels)
    av_floor_idx, av_floor_name = _floor_for_node_y(av_y_dxf, floor_labels)

    def _elev_for_node(ny: float) -> tuple[float, str | None, bool, int | None]:
        """노드 Y 의 (elev_m, floor_name, from_label, floor_idx). label 없으면 Y/1000 fallback."""
        if floor_labels and av_floor_idx is not None:
            f_idx, f_name = _floor_for_node_y(ny, floor_labels)
            if f_idx is not None:
                return ((f_idx - av_floor_idx) * floor_height_mm / 1000.0, f_name, True, f_idx)
        return ((ny - av_y_dxf) / 1000.0, None, False, None)

    # 노드 — 라벨 컨벤션:
    #   첫 노드 "1" (Input/펌프), 마지막 노드 "10" (AV).
    #   중간 노드는 "n2", "n3", ... ("10" 과 충돌 방지 — path 길이 ≥ 10 일 때 collision 버그 fix).
    nodes: list[dict] = []
    nodes_with_floor = 0
    for i, pt in enumerate(path):
        if i == 0:
            label, io = "1", "Input"
        elif i == total - 1:
            label, io = "10", "No"
        else:
            label, io = f"n{i + 1}", "No"
        elev_m, floor_name, from_label, floor_idx = _elev_for_node(pt[1])
        if from_label:
            nodes_with_floor += 1
        node: dict = {
            "label": label,
            "x": int(round(pt[0])),
            "y": int(round(pt[1])),
            "elevation": round(elev_m, 3),
            "io_node": io,
        }
        if floor_name:
            node["floor"] = floor_name
        if floor_idx is not None:
            node["floor_idx"] = floor_idx
        if io == "Input":
            node["pressure_pa"] = 101325.0
        nodes.append(node)

    # 파이프 + 직경 매칭
    pipes: list[dict] = []
    dia_match_count = 0
    for i in range(total - 1):
        a = path[i]; b = path[i + 1]
        edge_key = (min(a, b), max(a, b))
        length_mm = edge_len.get(edge_key, math.hypot(b[0] - a[0], b[1] - a[1]))
        length_m_dxf = length_mm / 1000.0
        dia, dia_dist, dia_raw = _match_diameter_for_segment(a, b, dia_text_pts)
        # C — 직경 default 100→150 (47 도면 학습:
        #    답안 main_bore 분포 100mm 165개 / 150mm 148개 거의 동률,
        #    대명동/양주옥정 자연낙차 case 답안 main 모두 150mm. 절충 default).
        used_dia = dia if dia is not None else 150
        if dia is not None:
            dia_match_count += 1
        # pipe elev: 노드 간 elev 차이 (층 라벨 매칭 시 floor-aware elev).
        in_e = nodes[i]["elevation"]
        out_e = nodes[i + 1]["elevation"]
        elev_m = round(out_e - in_e, 3)
        # PIPENET 제약: |elev| ≤ length (피타고라스). 도면이 짧게 압축돼 그려진
        # 수직 run (한 층 차이 = 2.1m elev 인데 DXF segment 가 1m 같은 경우) 의
        # 실제 길이는 elev 만큼 되어야 hydraulic 계산 가능. length 보정.
        length_m = max(length_m_dxf, abs(elev_m))
        total_length_mm += length_m * 1000.0
        # Pipe 라벨에 "r" prefix — 라이저(1..9)/헤드망(10+) 컨벤션 영역 분리.
        # path 길이 ≥ 10 이면 "10" 등이 헤드망 pipe 와 충돌 (stitch 시 ValueError).
        # "r1", "r2", ... 식으로 prefix 해 절대 겹칠 일 없게.
        pipe: dict = {
            "label": f"r{i + 1}",
            "in":  nodes[i]["label"],
            "out": nodes[i + 1]["label"],
            "type": "KSD 3507",
            "dia": used_dia,
            "length": round(length_m, 3),
            "elev":   elev_m,
            "c": "120",
            "status": "Normal",
            "group": "Unset",
        }
        if dia is not None:
            pipe["dia_source"] = "text_match"
            pipe["dia_match_dist_mm"] = round(dia_dist, 1)
            if dia_raw:
                pipe["dia_raw"] = dia_raw
        else:
            pipe["dia_source"] = "default"
        pipes.append(pipe)

    return {
        "nodes": nodes,
        "pipes": pipes,
        "pumps": [],
        "valves": [],
        "av_node_label": "10",
        "input_node_label": "1",
        "title": "SYSTEM_EXTRACT_V1",
        "zone_kind": "system_path_dxf",
        "extracted_from": "dxf",
        "path_node_count": total,
        "total_pipe_length_m": round(total_length_mm / 1000.0, 2),
        "pump_snap_dist_mm": round(pump_snap_dist, 1),
        "av_snap_dist_mm":   round(av_snap_dist, 1),
        "graph_stats": graph_stats or {},
        # v2 — 직경 / 층 매칭 통계
        "diameter_matching": {
            "matched_pipes": dia_match_count,
            "total_pipes":   len(pipes),
            "text_label_count": len(dia_text_pts),
        },
        "floor_matching": {
            "label_count": len(floor_labels),
            "floor_height_mm": round(floor_height_mm, 0),
            "av_floor_idx":  av_floor_idx,
            "av_floor_name": av_floor_name,
            "nodes_with_floor": nodes_with_floor,
        },
        # 호환성 키 — legacy template 출력 형태 유지
        "affine_scale": 1.0,
        "affine_rotation_deg": 0.0,
    }


def extract_clean_system_network(dxf_path, scale_mm_per_unit: float = 1.0) -> dict:
    """깨끗한(손작도) 배관망 DXF 의 **전체 망**을 그대로 riser dict 로.

    임시 stopgap — 풀 계통도가 조각나 강제 bridge 로 path 가 튀는 문제를 우회.
    깨끗한 파일(계통도_LH_306_배관망추출.dxf)은 단일 연결망이라 force_connect 없이
    파일에 그려진 선 그대로 배관 + 길이를 띄운다. 단일 P→AV path 가 아니라 망 전체를
    pipe 로 낸다.

    Args:
        scale_mm_per_unit: 도면 1단위 = 실제 몇 mm 인지. 1.0 이면 도면 측정값 그대로(용지
            스케일이면 작게 나옴). 실제 플롯 스케일을 알면 곱해서 실측 길이로 변환.
    """
    parsed = parse_dxf_for_view(dxf_path, include_hidden_layers=True)
    entities = parsed["entities"]
    if not entities:
        raise ValueError("배관망 DXF entity 비어있음")
    # 깨끗한 파일은 단일 컴포넌트 — force_connect 불필요(추측 bridge 0개라 안 튄다).
    graph, edge_len, stats = build_system_graph(entities, force_connect=False)
    if not graph:
        raise ValueError("LINE entity 가 없음 — 배관망 추출 불가")
    dia_text_pts = _extract_dia_text_points(entities)
    floor_labels = _extract_floor_labels(entities)
    return _network_to_riser_dict(
        graph, edge_len, stats=stats,
        dia_text_pts=dia_text_pts, floor_labels=floor_labels,
        scale_mm_per_unit=scale_mm_per_unit,
    )


def _network_to_riser_dict(
    graph: dict,
    edge_len: dict,
    stats: dict | None = None,
    dia_text_pts: list | None = None,
    floor_labels: list | None = None,
    scale_mm_per_unit: float = 1.0,
) -> dict:
    """그래프 전체(트리/망)를 riser dict 로 — 모든 edge 를 pipe 로, 측정 길이 포함.

    노드 라벨: degree-1 잎 하나를 "1"(Input), 그로부터 가장 먼 잎을 "10"(AV) 으로,
    나머지는 "n{i}". 단일 path 가 아니므로 분기(branch)도 그대로 pipe 로 낸다.
    """
    dia_text_pts = dia_text_pts or []
    floor_labels = floor_labels or []
    node_list = list(graph.keys())
    if not node_list:
        raise ValueError("배관망 노드가 없음")

    # ★ 다중 컴포넌트 정리 — 클린망에 떠 있는 작은 부유 조각(2~3 노드짜리 stray pipe)은
    # 평면도가 라이저로만 연결되거나 도면 노이즈라서, 단독 수리계산 시 솔버가 "특정 노드
    # 누락(연결 안 됨)" 으로 계산을 막는다(LH306동 평면: 628 본망 + 2노드 조각 6개 →
    # 7 컴포넌트로 SDF/KFP/HAS 전부 계산 불가 재현). 가장 큰 연결 컴포넌트만 남겨 단일망을
    # 보장한다. 이미 단일 컴포넌트면 무영향(no-op). 버려진 조각은 stats 에 기록.
    comps = _connected_components(graph)
    if len(comps) > 1:
        comps.sort(key=len, reverse=True)
        keep = comps[0]
        dropped = len(graph) - len(keep)
        graph = {n: [m for m in graph[n] if m in keep] for n in keep}
        edge_len = {e: L for e, L in edge_len.items() if e[0] in keep and e[1] in keep}
        node_list = list(graph.keys())
        if stats is not None:
            stats["dropped_fragment_components"] = len(comps) - 1
            stats["dropped_fragment_nodes"] = dropped
            stats["kept_component_nodes"] = len(keep)

    deg = {n: len(graph[n]) for n in node_list}
    leaves = [n for n in node_list if deg.get(n, 0) == 1]
    input_node = leaves[0] if leaves else node_list[0]

    # input 에서 가장 먼(그래프 거리) 잎 → AV 후보 (라벨 "10").
    av_node = None
    if len(node_list) > 1:
        far = None
        far_d = -1.0
        for leaf in (leaves or node_list):
            if leaf == input_node:
                continue
            p = _shortest_path(graph, edge_len, input_node, leaf)
            if not p:
                continue
            d = sum(
                edge_len.get((min(p[i], p[i + 1]), max(p[i], p[i + 1])),
                             math.hypot(p[i + 1][0] - p[i][0], p[i + 1][1] - p[i][1]))
                for i in range(len(p) - 1)
            )
            if d > far_d:
                far_d, far = d, leaf
        av_node = far

    ref_y = input_node[1]
    label_of: dict = {}
    nodes: list[dict] = []
    ni = 1
    for n in node_list:
        if n == input_node:
            label, io = "1", "Input"
        elif n == av_node:
            label, io = "10", "No"
        else:
            ni += 1
            label, io = f"n{ni}", "No"
        label_of[n] = label
        elev_m = (n[1] - ref_y) / 1000.0 * scale_mm_per_unit
        node: dict = {
            "label": label,
            "x": int(round(n[0])),
            "y": int(round(n[1])),
            "elevation": round(elev_m, 3),
            "io_node": io,
        }
        if io == "Input":
            node["pressure_pa"] = 101325.0
        nodes.append(node)

    pipes: list[dict] = []
    seen: set = set()
    total_length_mm = 0.0
    total_measured_mm = 0.0
    dia_match_count = 0
    pidx = 0
    for a in graph:
        for b in graph[a]:
            key = (a, b) if a <= b else (b, a)
            if key in seen:
                continue
            seen.add(key)
            measured_mm = edge_len.get(key, math.hypot(b[0] - a[0], b[1] - a[1]))
            total_measured_mm += measured_mm
            length_m = measured_mm / 1000.0 * scale_mm_per_unit
            dia, dia_dist, dia_raw = _match_diameter_for_segment(a, b, dia_text_pts)
            used_dia = dia if dia is not None else 150
            if dia is not None:
                dia_match_count += 1
            in_e = (a[1] - ref_y) / 1000.0 * scale_mm_per_unit
            out_e = (b[1] - ref_y) / 1000.0 * scale_mm_per_unit
            elev_m = round(out_e - in_e, 3)
            length_m = max(length_m, abs(elev_m))
            total_length_mm += length_m * 1000.0
            pidx += 1
            pipe: dict = {
                "label": f"r{pidx}",
                "in": label_of[a],
                "out": label_of[b],
                "type": "KSD 3507",
                "dia": used_dia,
                "length": round(length_m, 3),
                "length_measured_mm": round(measured_mm, 2),
                "elev": elev_m,
                "c": "120",
                "status": "Normal",
                "group": "Unset",
                "dia_source": "text_match" if dia is not None else "default",
            }
            pipes.append(pipe)

    return {
        "nodes": nodes,
        "pipes": pipes,
        "pumps": [],
        "valves": [],
        "av_node_label": "10" if av_node is not None else None,
        "input_node_label": "1",
        "title": "CLEAN_NETWORK (임시)",
        "zone_kind": "clean_network_dxf",
        "extracted_from": "dxf_clean_network",
        "path_node_count": len(nodes),
        "total_pipe_length_m": round(total_length_mm / 1000.0, 2),
        "total_measured_mm": round(total_measured_mm, 1),
        "scale_mm_per_unit": scale_mm_per_unit,
        "network_edges": _graph_edges_for_render(graph),
        "graph_stats": stats or {},
        "diameter_matching": {
            "matched_pipes": dia_match_count,
            "total_pipes": len(pipes),
            "text_label_count": len(dia_text_pts),
        },
        "affine_scale": 1.0,
        "affine_rotation_deg": 0.0,
    }


# ────────────────────────────────────────────────────────────────────────────
# 기계실(옥상수조) 경로 추출 — extract_system_path 미러
# ────────────────────────────────────────────────────────────────────────────

# 기계실 소화배관 레이어 (대명동 201동 옥상층 평면도 기준).
# 스프링클러(SP) 계통만 — 물탱크 레이어는 배관이 아니라 탱크 박스 도면(외곽선·
# 해치 LINE 54+)이라 그래프에 넣으면 가짜 배관 노드 60+ 와 추정 bridge 다발이
# 생겨 평면도가 엉키고 배치 bbox 가 틀어진다. 탱크는 수원 한 점(m1)으로만 표현하며,
# SP 배관이 탱크 토출구 근처(실측 252mm)에 끝점을 가지므로 수원 스냅은 그대로 동작.
# 도면별 레이어명이 달라 layer_filter 미지정 시 존재하는 것만 추려 쓰고,
# 매칭 0건이면 build_system_graph 의 키워드 자동 필터 사용.


def extract_machine_room_path(
    entities: list[dict],
    source_xy: tuple[float, float],
    riser_conn_xy: tuple[float, float],
    snap_tolerance_mm: float = 2500.0,
    layer_filter: set[str] | None = None,
) -> dict:
    """기계실(옥상수조) DXF 에서 수원(탱크) → 입상관 연결점 배관 경로 추출.

    계통도 추출(extract_system_path)과 동형 — 같은 그래프/Dijkstra 사용.
    차이는 (1) 의미: source = 옥상수조 수원(Input 경계), conn = 라이저 Input '1'
    에 stitch 될 입상관 연결점. (2) 라벨: m1..mK ('1'~'10'/'r*'/헤드 10+ 와 비충돌).

    Args:
        entities: parse_dxf_for_view().entities.
        source_xy: 사용자 픽 탱크 토출구(수원) 좌표 (mm).
        riser_conn_xy: 사용자 픽 입상관 연결점 좌표 (mm).
        snap_tolerance_mm: 클릭 ↔ 그래프 노드 허용 거리.
        layer_filter: None 이면 SP+물탱크 레이어 시도, 결과 부족 시 키워드 자동.

    Returns:
        { nodes, pipes, source_node_label, conn_node_label, ... 진단 }
        nodes[0] = m1 (source, io='Input', pressure_pa=101325),
        nodes[-1] = mK (conn, io='No') — 3-way stitch 시 라이저 '1' 과 병합.

    Raises:
        ValueError: snap 실패 / 경로 없음.
    """
    if not entities:
        raise ValueError("기계실 entity 비어있음 — DXF 파싱 결과 확인 필요")

    lf = layer_filter
    if lf is None:
        present = {en.get("l") for en in entities}
        lf = {ly for ly in MACHINE_ROOM_SP_LAYERS if ly in present} or None

    # force_connect(무제한 봉합)는 기계실 추출 전용 — anchored 평면도 경로에서는
    # 호출 금지 (W3: 표적 브릿지로 대체).
    graph, edge_len, stats = build_system_graph(entities, layer_filter=lf,
                                                force_connect=True)
    if not graph:
        raise ValueError(f"LINE entity 없음 (전체 {len(entities)}개 중 LINE/PL 0개)")

    # 클릭 → 가장 가까운 그래프 노드. 끝점 거리 제한 없이 무조건 가장 가까운 노드로
    # 끌어붙인다 (사용자 요구: "끝점 제한 없애고, pipe_소화배관만 켜도 그 위에서 최소거리
    # 그냥 따라가게"). 선택 레이어(lf)로 그래프가 그 배관만 담고, force_connect=True 로
    # 단일 컴포넌트라 어떤 클릭이든 그 배관 위 최근접 노드로 snap → 실패(ValueError) 없음.
    # snap_tolerance_mm 인자는 더 이상 거리 컷에 쓰지 않는다(진단/호환용으로만 유지).
    def _snap(xy):
        n = _nearest_graph_node(graph, (float(xy[0]), float(xy[1])))
        d = math.hypot(n[0] - float(xy[0]), n[1] - float(xy[1])) if n else float("inf")
        return n, d

    src_node, src_d = _snap(source_xy)
    conn_node, conn_d = _snap(riser_conn_xy)
    if src_node is None or conn_node is None:
        raise ValueError("그래프에 노드가 없어 수원/연결점을 매핑할 수 없음 "
                         "(선택 배관 레이어에 LINE/PL 이 있는지 확인).")

    # force_connect 가 거리 무제한으로 강제 연결한 bridge(추정 연결)는 실제 배관이
    # 아니다. 이를 실선으로 같이 그리면 도면 전체를 가로지르는 직선들이 생겨 "꼬여
    # 보이는" 원인이 된다. → 실측 edge(plan_edges)와 추정 edge(plan_edges_estimated)를
    # 분리하고, 추출 경로(spine)가 추정 bridge 를 통과하면 그 segment 를 표시해
    # 프론트가 점선·경고색으로 렌더하게 한다 (= "이 연결은 도면에 없는 알고리즘 추정").
    forced_keys: set[tuple] = set()
    for (ea, eb) in (stats.get("forced_bridge_edges") or []):
        ka = (int(ea[0]), int(ea[1]))
        kb = (int(eb[0]), int(eb[1]))
        forced_keys.add((min(ka, kb), max(ka, kb)))

    # 추정 bridge 는 거대 가중치로 패널티 → Dijkstra 가 실배관 경로를 우선한다.
    # (추정 bridge 의 가중치는 실제 직선거리라 가장 짧아 패널티 없이는 wormhole 처럼
    #  선호돼 "엉뚱한 경로"가 된다.) 진짜 다른 길이 없을 때만 최소 개수로 사용.
    path = _shortest_path(graph, edge_len, src_node, conn_node, penalty_keys=forced_keys)
    if not path or len(path) < 2:
        raise ValueError(
            f"수원 → 입상관 연결점 경로 없음 — disconnected component 가능. "
            f"그래프 컴포넌트 {stats['components_after_bridge']}개 "
            f"(bridge {stats['bridges_applied']}회 후). snap 수원={src_d:.0f}mm, 연결={conn_d:.0f}mm."
        )

    path = _collapse_collinear_nodes(path, edge_len, angle_tol_deg=2.0)
    dia_text_pts = _extract_dia_text_points(entities)

    result = _machine_room_path_to_dict(
        path, edge_len, source_xy, riser_conn_xy,
        source_snap_dist=src_d, conn_snap_dist=conn_d,
        graph_stats=stats, dia_text_pts=dia_text_pts,
        forced_keys=forced_keys,
    )

    # 전체 SP 배관망 edge (시각화 전용) — 수리경로 spine 뿐 아니라 기계실 전 배관을
    # 평면도로 렌더하기 위해 그래프 전 edge 를 raw DXF 좌표로 dedupe 해 첨부.
    seen: set[tuple] = set()
    plan_edges: list[list[float]] = []
    plan_edges_estimated: list[list[float]] = []
    for a, nbrs in graph.items():
        for b in nbrs:
            key = (min(a, b), max(a, b))
            if key in seen:
                continue
            seen.add(key)
            ra = (int(round(a[0])), int(round(a[1])))
            rb = (int(round(b[0])), int(round(b[1])))
            rounded_key = (min(ra, rb), max(ra, rb))
            edge = [ra[0], ra[1], rb[0], rb[1]]
            if rounded_key in forced_keys:
                plan_edges_estimated.append(edge)
            else:
                plan_edges.append(edge)
    result["plan_edges"] = plan_edges
    result["plan_edges_estimated"] = plan_edges_estimated
    return result


def _machine_room_path_to_dict(
    path: list[tuple[float, float]],
    edge_len: dict,
    source_xy_orig: tuple[float, float],
    conn_xy_orig: tuple[float, float],
    source_snap_dist: float = 0.0,
    conn_snap_dist: float = 0.0,
    graph_stats: dict | None = None,
    dia_text_pts: list[tuple[float, float, int, str]] | None = None,
    forced_keys: set[tuple] | None = None,
) -> dict:
    """기계실 경로(vertex 시퀀스) → dict. 라벨 m1..mK, m1=Input(옥상수조 수면, 1atm).

    옥상수조부는 수평 분포라 노드 elev=0 (탱크 수면 기준). 실제 수직 낙차는
    라이저(계통도)가 담당하므로 기계실 elev 는 의도적으로 0 으로 둔다.
    """
    if len(path) < 2:
        raise ValueError(f"경로 노드 수 {len(path)} — 수원 = 연결점 같은 위치 가능성")
    dia_text_pts = dia_text_pts or []
    total = len(path)
    total_length_mm = 0.0

    nodes: list[dict] = []
    for i, pt in enumerate(path):
        io = "Input" if i == 0 else "No"
        node: dict = {
            "label": f"m{i + 1}",
            "x": int(round(pt[0])),
            "y": int(round(pt[1])),
            "elevation": 0.0,
            "io_node": io,
        }
        if io == "Input":
            node["pressure_pa"] = 101325.0  # 개방형 옥상수조 수면 = 대기압 경계
        nodes.append(node)

    forced_keys = forced_keys or set()
    estimated_seg_count = 0
    pipes: list[dict] = []
    dia_match_count = 0
    for i in range(total - 1):
        a = path[i]; b = path[i + 1]
        edge_key = (min(a, b), max(a, b))
        length_mm = edge_len.get(edge_key, math.hypot(b[0] - a[0], b[1] - a[1]))
        total_length_mm += length_mm
        ra = (int(round(a[0])), int(round(a[1])))
        rb = (int(round(b[0])), int(round(b[1])))
        is_estimated = (min(ra, rb), max(ra, rb)) in forced_keys
        if is_estimated:
            estimated_seg_count += 1
        dia, dia_dist, dia_raw = _match_diameter_for_segment(a, b, dia_text_pts)
        used_dia = dia if dia is not None else 150
        if dia is not None:
            dia_match_count += 1
        pipe: dict = {
            "label": f"m{i + 1}",
            "in":  nodes[i]["label"],
            "out": nodes[i + 1]["label"],
            "type": "KSD 3507",
            "dia": used_dia,
            "length": round(length_mm / 1000.0, 3),
            "elev": 0.0,
            "c": "120",
            "status": "Normal",
            "group": "Unset",
        }
        if dia is not None:
            pipe["dia_source"] = "text_match"
            pipe["dia_match_dist_mm"] = round(dia_dist, 1)
            if dia_raw:
                pipe["dia_raw"] = dia_raw
        else:
            pipe["dia_source"] = "default"
        if is_estimated:
            # 도면에 없는 알고리즘 추정 연결 — 프론트가 점선·경고색으로 표시.
            pipe["estimated"] = True
        pipes.append(pipe)

    return {
        "nodes": nodes,
        "pipes": pipes,
        "estimated_segment_count": estimated_seg_count,
        "source_node_label": "m1",
        "conn_node_label": f"m{total}",
        "title": "MACHINE_ROOM_EXTRACT_V1",
        "zone_kind": "machine_room_path_dxf",
        "extracted_from": "dxf",
        "path_node_count": total,
        "total_pipe_length_m": round(total_length_mm / 1000.0, 2),
        "source_snap_dist_mm": round(source_snap_dist, 1),
        "conn_snap_dist_mm": round(conn_snap_dist, 1),
        "graph_stats": graph_stats or {},
        "diameter_matching": {
            "matched_pipes": dia_match_count,
            "total_pipes": len(pipes),
            "text_label_count": len(dia_text_pts),
        },
    }


def filter_pipenet_only(bundle: ParsedDxfBundle) -> list[dict]:
    """Stage 1 — 배관망 관련 entity 만 필터 (auto_category in PIPE/HEAD/TEXT or layer in KEEP_BASE_LAYERS)."""
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    out = []
    for en in bundle.entities:
        cat = layer_cat.get(en["l"], "OTHER")
        if cat in PIPENET_CATEGORIES or en["l"] in KEEP_BASE_LAYERS:
            out.append(en)
    return out


# ────────────────────────────────────────────────────────────────────────────
# 2) Stage 2 — G₀ 그래프 빌드 + 가장 불리한 K 헤드 + subgraph 추출
# ────────────────────────────────────────────────────────────────────────────

# 50mm: DN50 (최소 호칭경) 미만 거리는 배관 토폴로지상 같은 점으로 간주.
# 부동소수점 오차 + DWG→DXF 변환 누적 오차 + CAD 작업자 미세 오차 모두 흡수.
# 이전 5mm 는 대명동/다이소 작업엔 문제 없었으나, 좌표 절댓값이 큰 도면
# (예: MF-125 의 측지 좌표 3,500,000mm) 에서 변환 오차가 5mm 초과 → SP-LINE
# 끝점들이 안 만나서 그래프가 3,058 component 로 쪼개지는 사고 발생.
# 50mm 는 토폴로지 분석에 영향 없음 (호칭경 단위가 50/65/80/100mm 라 50mm
# 이내 차이는 의미 없음). 격자 snap 아니라 _NodeIndex cluster 반경.
# 5m: 메자닌/대형 도면 (예: MF-125) 의 헤드가 배관 라인과 천장고 차이로 멀리 떨어진 경우 보호.
# 알람밸브는 라이저 (수직 입상관) 위에 위치 — 평면도상 가지관과 거리가 멀 수 있음.
# 25m 이내면 알람밸브 위치 그대로 source 로 사용 → 그래프 component 통합 효과.
ESTIMATED_BRIDGE_MM = 2000.0  # 이 길이 초과 강제 bridge = 추정(도면에 없는 wormhole).
# 2m 이하 gap 은 fitting/도면오차로 보고 실배관 연속으로 취급(경로에 사용). 초과분은 추정
# bridge(노란 점선)로 보고 Dijkstra penalty → 실배관 대안이 있으면 망 생성에서 회피.
# 50mm 미만 LINE/PL/ARC segment 는 그래프 edge 로 사용 안 함.
# 헤드 부속(HEADCON, HDCROSS, SPCAP 등), 치수 보조선, 텍스트 underline 등
# 평면도에는 보이지만 배관망 토폴로지에는 노이즈인 짧은 segment 제거.






@dataclass(slots=True)
class HeadCandidate:
    pos: tuple[float, float]  # snapped
    raw: tuple[float, float]  # original coord (for SDF Position)
    block_name: str
    layer: str


# 참조 5종 head DXF 분석에서 얻은 알려진 블록 이름 (modelspace 직접 INSERT 또는 nested)
# 사용자가 업로드한 헤드 DXF 들의 BLOCKS section 정의:
#   A$C39172136 — 폐쇄형 SP-HEAD (메인, 대명동 도면 111회 사용)
#   A$C3F157AFD — 조기반응형 폐쇄형 105도
#   A$C60792707 — 조기반응형 폐쇄형 72도
#   A$C6B5253FE — head nested (depth 2)
#   A$C563427C5 — head nested (depth 3)
#   A$C324C7814 — head body block (LWPOLYLINE + CIRCLE)
#   A$C0F5C7CDB — head fitting (LWPOLYLINE + CIRCLE x 2)
KNOWN_HEAD_BLOCKS: set[str] = {
    # 대명동 201동 (기존 — nested INSERT 블록명)
    "A$C39172136", "A$C3F157AFD", "A$C60792707",
    "A$C6B5253FE", "A$C563427C5", "A$C324C7814", "A$C0F5C7CDB",
    # 다이소 양주허브센터 — 47 도면 분석 결과 발견 (총 34K+ 인스턴스)
    "K-160 헤드",            # 18,523 — K-160 표준 스프링클러
    "K-160 (조기반응)",       #  9,042 — 조기반응형 스프링클러
    "K-200 헤드",            #  7,039 — K-200 큰 직경
    "Large Drop head-1",     #  1,665 — 다이소 물류센터 Large Drop 헤드
    # 양주옥정 중상1블럭 — 표준 SP01 시리즈 (총 63K+ 인스턴스)
    "SP01-01",               #  1,138
    "SP01-02",               #  6,031
    "SP01-04",               # 54,837 — 표준 헤드
    "SP01-05",               #  2,194
}

# 헤드 부속/연결 (헤드 자체 아니지만 헤드 근처에 같이 작도되는 블록).
# 이 블록들은 헤드 근접 클러스터링에 가중치 추가용 — 단독으로는 헤드 안 됨.
HEAD_FITTING_BLOCKS: set[str] = {
    "HEADCON",      # 15,193 — 헤드 connection (배관 연결)
    "HDCROSS",      # 12,117 — 헤드 cross fitting (T자)
    "HEADCOL",      #  4,204 — 헤드 collar
    "HEADCOR",      #  3,980 — 헤드 corner
    "SPCAP",        # 32,643 — SP 캡 (마감)
    "HEADCOL (3)",  #    883 — collar 3-way
}


@dataclass(slots=True)
class HeadDetection:
    """전체 헤드 인식 결과 — 도면 내 한 헤드의 바운딩박스 + 메타."""

    pos: tuple[float, float]               # 헤드 중심 (world coord)
    bbox: tuple[float, float, float, float]  # x1, y1, x2, y2 (world coord)
    kind: str                              # 인식 방법 (block_match / circle_signature / hatch_triangle / cluster)
    confidence: float                      # 0~1
    block_name: str = ""
    layer: str = ""




def detect_heads(pipe_entities: list[dict], layer_categories: dict[str, str],
                 region=None) -> list[HeadDetection]:
    """도면 내 모든 헤드 후보 인식 — 다중 신호 결합 + 근접 클러스터링.

    인식 규칙
    ---------
    R1) HEAD 카테고리 레이어의 INSERT — block name 이 KNOWN_HEAD_BLOCKS 면 confidence 0.95,
        그 외 HEAD layer INSERT 는 0.70
    R2) HEAD 카테고리 레이어의 CIRCLE 중 반경 10~250mm — confidence 0.80 (head 본체 마커)
    R3) HEAD 카테고리 레이어의 HATCH (드라이팬던트 삼각형 등) — confidence 0.75
    R5) **layer-agnostic 삼각형 HATCH** — 3 고유 정점 + bbox < 1500mm 면 confidence 0.72
        (드라이팬던트 헤드 마커 — 참조 elbow/측벽 DXF 처럼 HEAD 레이어 아닌 곳도 검출)
    R4) 클러스터링 — 250mm 이내 후보들을 1 헤드로 통합 (cue 가 여러 개일수록 confidence ↑)

    region: anchored 모드의 헤드 영역 게이트. ``contains((x, y)) -> bool`` 프로토콜
        객체(→W4 HeadRegion). 지정 시 **최종 승인 후보에만** point-in-region 판정을
        적용해 범례 표본·인접 세대 헤드를 배제한다. R1~R5 신호 계산·클러스터링은
        불변이며 ``region=None`` 이면 기존과 완전 동일.
    """
    candidates: list[HeadDetection] = []

    for en in pipe_entities:
        cat = layer_categories.get(en.get("l", ""), "OTHER")
        # R1/R2/R3 — HEAD 카테고리 전용
        if cat == "HEAD":
            if en["t"] == "I":
                x, y = float(en["p"][0]), float(en["p"][1])
                bn = en.get("n", "")
                conf, kind = (0.95, "block_match") if bn in KNOWN_HEAD_BLOCKS else (0.70, "head_layer_insert")
                bbox = (x - 100.0, y - 100.0, x + 100.0, y + 100.0)
                candidates.append(HeadDetection(pos=(x, y), bbox=bbox, kind=kind,
                                                confidence=conf, block_name=bn, layer=en["l"]))
            elif en["t"] == "C":
                cx, cy = float(en["c"][0]), float(en["c"][1])
                r = float(en.get("r", 0))
                if 10.0 <= r <= 250.0:
                    bbox = (cx - r - 30, cy - r - 30, cx + r + 30, cy + r + 30)
                    candidates.append(HeadDetection(pos=(cx, cy), bbox=bbox, kind="circle_signature",
                                                    confidence=0.80, layer=en["l"]))
            elif en["t"] == "H":
                pts = en.get("p", [])
                if len(pts) >= 3:
                    xs = [float(p[0]) for p in pts]
                    ys = [float(p[1]) for p in pts]
                    w = max(xs) - min(xs); h = max(ys) - min(ys)
                    if w <= 1500 and h <= 1500:
                        cx = sum(xs) / len(xs); cy = sum(ys) / len(ys)
                        bbox = (min(xs) - 20, min(ys) - 20, max(xs) + 20, max(ys) + 20)
                        candidates.append(HeadDetection(pos=(cx, cy), bbox=bbox, kind="hatch_triangle",
                                                        confidence=0.75, layer=en["l"]))

        # R5 — layer-agnostic 삼각형 HATCH (드라이팬던트 헤드)
        # HEAD 카테고리 아닌 곳도 검사. 단, 정확히 3 고유 정점 + bbox ≤ 1500mm 일 때만.
        if en["t"] == "H" and cat != "HEAD":
            pts = en.get("p", [])
            if _is_triangle_shape(pts):
                xs = [float(p[0]) for p in pts]
                ys = [float(p[1]) for p in pts]
                w = max(xs) - min(xs); h = max(ys) - min(ys)
                if w <= 1500 and h <= 1500:
                    cx = sum(xs) / len(xs); cy = sum(ys) / len(ys)
                    bbox = (min(xs) - 20, min(ys) - 20, max(xs) + 20, max(ys) + 20)
                    candidates.append(HeadDetection(
                        pos=(cx, cy), bbox=bbox,
                        kind="triangle_drypendant", confidence=0.72, layer=en["l"],
                    ))

    # ── 클러스터링 — 같은 헤드를 가리키는 여러 cue (INSERT + CIRCLE + HATCH) 를 한 개로 ──
    # 공간 격자(grid) 인덱스로 근접 후보만 비교 → 전수 비교 O(N²) 제거. seed(가장
    # 앞선 미사용 후보) 기준 반경 클러스터링이라, 셀 크기를 CLUSTER_R 로 잡으면 반경
    # 내 후보는 항상 seed 셀의 3×3 이웃 안에 있다. 클러스터 집계(max/min/set/len)는
    # 순서 무관이라 결과는 전수 비교와 동일.
    CLUSTER_R = 250.0
    grid: dict[tuple[int, int], list[int]] = {}
    for idx, c in enumerate(candidates):
        key = (int(math.floor(c.pos[0] / CLUSTER_R)), int(math.floor(c.pos[1] / CLUSTER_R)))
        grid.setdefault(key, []).append(idx)

    used = [False] * len(candidates)
    out: list[HeadDetection] = []
    for i, c1 in enumerate(candidates):
        if used[i]:
            continue
        cluster = [c1]
        used[i] = True
        gx0 = int(math.floor(c1.pos[0] / CLUSTER_R))
        gy0 = int(math.floor(c1.pos[1] / CLUSTER_R))
        for gx in (gx0 - 1, gx0, gx0 + 1):
            for gy in (gy0 - 1, gy0, gy0 + 1):
                for j in grid.get((gx, gy), ()):
                    if j <= i or used[j]:
                        continue
                    c2 = candidates[j]
                    if math.hypot(c1.pos[0] - c2.pos[0], c1.pos[1] - c2.pos[1]) <= CLUSTER_R:
                        cluster.append(c2)
                        used[j] = True
        best = max(cluster, key=lambda c: c.confidence)
        x1 = min(c.bbox[0] for c in cluster)
        y1 = min(c.bbox[1] for c in cluster)
        x2 = max(c.bbox[2] for c in cluster)
        y2 = max(c.bbox[3] for c in cluster)
        # 클러스터 cue 가 많을수록 confidence ↑ (최대 0.99)
        conf = min(0.99, best.confidence + 0.05 * (len(cluster) - 1))
        kinds = "+".join(sorted({c.kind for c in cluster}))
        out.append(HeadDetection(
            pos=best.pos, bbox=(x1, y1, x2, y2),
            kind=kinds if len(cluster) == 1 else f"cluster({len(cluster)}):{kinds}",
            confidence=conf, block_name=best.block_name, layer=best.layer,
        ))
    if region is not None:
        # anchored 영역 게이트 — 클러스터링까지 끝난 최종 승인 후보에만 적용.
        out = [h for h in out if region.contains(h.pos)]
    return out


def find_unreachable_region_heads(
    graph: dict,
    source: tuple[float, float] | None,
    heads: list[HeadDetection],
    max_attach_mm: float = HEAD_BRIDGE_MAX_MM,
) -> list[tuple[float, float]]:
    """region 승인 헤드 중 source 에서 도달 불가한 헤드 좌표 목록.

    anchored 모드 진단 — 조용한 drop 금지: 미도달 헤드는 추출 결함 신호라
    audit(→W7)에 기록한다. 판정은 파이프라인의 헤드 부착 규칙과 동일하게
    최근접 그래프 노드(HEAD_BRIDGE_MAX_MM 이내)를 쓰되, 그 노드가 source 와
    같은 연결 컴포넌트가 아니면 미도달로 본다.
    """
    if source is None or source not in graph:
        return [h.pos for h in heads]
    comp = {source}
    stack = [source]
    while stack:
        n = stack.pop()
        for m in graph.get(n, ()):
            if m not in comp:
                comp.add(m)
                stack.append(m)
    unreachable: list[tuple[float, float]] = []
    for h in heads:
        near = _nearest_graph_node(graph, h.pos)
        if near is None:
            unreachable.append(h.pos)
            continue
        d = math.hypot(h.pos[0] - near[0], h.pos[1] - near[1])
        if d > max_attach_mm or near not in comp:
            unreachable.append(h.pos)
    return unreachable


def attach_source(
    alarm_xy: tuple[float, float],
    graph: dict,
    comp_of: dict,
    accepted_heads: list,
    edge_len: dict,
    audit: dict,
) -> tuple[tuple[float, float], tuple | None]:
    """anchored 소스 결합(W2) — blind ``_nearest_graph_node`` 단독 폴백 금지.

    1순위 후보 = region 내 승인 헤드를 1개 이상 보유한 컴포넌트의 노드들.
    그중 최근접에 부착. ``SOURCE_BRIDGE_MAX_MM`` 내에 없으면 무헤드 컴포넌트까지
    1단계 완화(escalation=1)하되 거리 상한은 유지. 그래도 없으면 anchored 실패 —
    명시적 에러(ValueError). 선택 근거(거리·컴포넌트 헤드 수·완화 단계)는
    ``audit['source_attach']`` 에 기록한다.

    accepted_heads: HeadDetection 목록(또는 (x, y) 좌표 목록).
    반환: (source 노드, 접속 edge 키 또는 None). alarm_xy 가 기존 노드와 1e-3 이상
    떨어져 있으면 alarm_xy 를 그래프 노드로 추가하고 최근접 후보 노드와 edge 로
    잇는다(기존 canonical 부착과 동일 방식) — 이 edge 는 추정연결이라 호출부에서
    라우팅 penalty 대상으로 등록해야 한다.
    """
    ax, ay = float(alarm_xy[0]), float(alarm_xy[1])
    head_count: dict = {}
    for h in accepted_heads:
        pos = h.pos if hasattr(h, "pos") else (float(h[0]), float(h[1]))
        near = _nearest_graph_node(graph, pos)
        if near is None:
            continue
        if math.hypot(pos[0] - near[0], pos[1] - near[1]) <= HEAD_BRIDGE_MAX_MM:
            cid = comp_of.get(near)
            if cid is not None:
                head_count[cid] = head_count.get(cid, 0) + 1

    def _nearest_in(pred) -> tuple:
        best, best_d = None, float("inf")
        for n in graph:
            if not pred(comp_of.get(n)):
                continue
            d = math.hypot(n[0] - ax, n[1] - ay)
            if d < best_d:
                best, best_d = n, d
        return best, best_d

    stages = (
        (0, "head_component_nearest", lambda cid: cid in head_count),
        (1, "any_component_nearest", lambda cid: True),
    )
    for escalation, method, pred in stages:
        node, d = _nearest_in(pred)
        if node is None or d > SOURCE_BRIDGE_MAX_MM:
            continue
        audit["source_attach"] = {
            "dist_mm": d,
            "method": method,
            "escalation": escalation,
            "comp_head_count": head_count.get(comp_of.get(node), 0),
        }
        if d <= 1e-3:
            return node, None
        src = (ax, ay)
        graph.setdefault(src, set()).add(node)
        graph[node].add(src)
        key = (min(src, node), max(src, node))
        edge_len[key] = d
        return src, key
    audit["source_attach"] = {
        "dist_mm": None, "method": "failed", "escalation": len(stages),
        "comp_head_count": 0,
    }
    raise ValueError(
        f"anchored 소스 결합 실패 — alarm_xy=({ax}, {ay}) 반경 "
        f"{SOURCE_BRIDGE_MAX_MM:.0f}mm 안에 부착 가능한 컴포넌트 없음"
    )


@dataclass(slots=True)
class SelectionResult:
    source_pos: tuple[float, float] | None
    source_kind: str
    heads: list[HeadCandidate]
    distances: list[float]
    edges: list[tuple[tuple[float, float], tuple[float, float], float]]  # merged pipes (a, b, length_mm)
    nodes_in_subgraph: list[tuple[float, float]]
    # 추가: pipe-내부에 흡수된 elbow 들. {(a,b): [(node_pos, angle_deg), ...]}
    elbow_fittings: dict[tuple, list[tuple[tuple[float, float], float]]] = field(default_factory=dict)
    # source 가 그래프 nearest 와 떨어진 거리(mm). 0 = 그래프 위에 정확히 있음 / 큰 값 = 떨어져 있음.
    source_bridge_dist_mm: float = 0.0
    # 한도(SOURCE_BRIDGE_MAX_MM) 초과로 source 를 nearest 로 fallback 한 경우 True.
    source_fallback: bool = False
    # anchored 실행의 추출 근거 리포트(W7). 비-anchored 경로에선 항상 None.
    audit: "ExtractionAudit | None" = None


@dataclass(slots=True)
class ExtractionAudit:
    """anchored 추출 근거 리포트(W7) — 파이프라인 audit 축적 dict 의 정식 스키마.

    기존 자료구조(weld/bridge/head-drop edge 키 set, attach_source·bridge_targeted
    의 audit 기록)를 그대로 재사용해 채운다 — 중복 계산 금지.
    """
    heads: dict = field(default_factory=dict)          # {detected_in_region, attached, unreachable:[[x,y],...]}
    bridges: list = field(default_factory=list)        # [{p1,p2,len_mm,tol,layers,p1_in_source_comp}]
    welds: list = field(default_factory=list)          # [{p1,p2,len_mm}]
    head_drops: list = field(default_factory=list)     # [{p1,p2,len_mm}]
    nonnominal: dict = field(default_factory=dict)     # {edge_count,len_mm,ratio}
    corridor: dict = field(default_factory=dict)       # {node_count,len_mm}
    source_attach: dict = field(default_factory=dict)  # {dist_mm,method,escalation,...}
    # 급수 감사 — {dead_edge_count,dead_len_mm,dead_ratio,watered_edge_count,heads_routed}
    water: dict = field(default_factory=dict)
    # T분기 edge-split 근거. sym_mm 은 "접속=기호" 관측 증거일 뿐 판정에 쓰지 않는다.
    tee_splits: list = field(default_factory=list)  # [{p,edge,gap_mm,sym_mm}]

    @classmethod
    def from_audit_dict(cls, audit: dict) -> "ExtractionAudit":
        """축적 dict → 스키마 정규화. anchor_window(W6, 객체) 등 비직렬화 항목 제외."""
        return cls(
            heads=dict(audit.get("heads") or {}),
            bridges=list(audit.get("bridges") or []),
            welds=list(audit.get("welds") or []),
            head_drops=list(audit.get("head_drops") or []),
            nonnominal=dict(audit.get("nonnominal") or
                            {"edge_count": 0, "len_mm": 0.0, "ratio": 0.0}),
            corridor=dict(audit.get("corridor") or {"node_count": 0, "len_mm": 0.0}),
            source_attach=dict(audit.get("source_attach") or {}),
            water=dict(audit.get("water") or
                       {"dead_edge_count": 0, "dead_len_mm": 0.0, "dead_ratio": 0.0,
                        "watered_edge_count": 0, "heads_routed": 0}),
            tee_splits=list(audit.get("tee_splits") or []),
        )

    def to_json_dict(self) -> dict:
        return {
            "heads": self.heads, "bridges": self.bridges, "welds": self.welds,
            "head_drops": self.head_drops, "nonnominal": self.nonnominal,
            "corridor": self.corridor, "source_attach": self.source_attach,
            "water": self.water, "tee_splits": self.tee_splits,
        }


def _build_graph(
    pipe_entities: list[dict],
    node_index: _NodeIndex | None = None,
    layer_categories: dict[str, str] | None = None,
    min_edge_mm: float = MIN_PIPE_EDGE_MM,
) -> tuple[dict[tuple[float, float], set[tuple[float, float]]], dict[tuple, float]]:
    """파이프 LINE/PL/ARC 으로부터 무방향 그래프 빌드.

    노이즈 컷:
      - layer_categories 가 주어지면 "PIPE" 카테고리 layer 의 entity 만 사용
        (헤드 부속 LINE, 텍스트 underline, 치수 보조선 등 제외)
      - closed PL (첫점=끝점) 은 배관 아니므로 제외 (알람밸브 박스 등)
      - min_edge_mm 미만 segment 는 그래프 edge 로 사용 안 함

    Geometry fallback (2026-06-22):
      배관이 전용 레이어가 아니라 기본 레이어 '0'(또는 이름에 키워드 없는 레이어)에
      작도된 도면(예: LH306동)은 "PIPE" 카테고리 필터가 전부 걸러 그래프가 빈다.
      → PIPE-strict 로 먼저 시도하고 edge 가 0 이면 HEAD/TEXT/ALARM/ARCH/EXCLUDE 를
      제외한 나머지(OTHER·기본 '0' 포함)를 배관 geometry 로 재시도. 이름이 제대로 된
      도면은 1차 통과로 동일 동작, 키워드 미스인 도면만 fallback 으로 살린다.

    Endpoint 동등성: _NodeIndex (epsilon=SNAP_TOL_MM mm) 기반 cluster.
    노드 좌표는 raw (DXF 원본) — 격자에 정렬 안 됨, 시각화 시 비뚤어짐 없음.

    node_index: caller 가 이미 가지고 있으면 재사용 (헤드/AV 좌표도 동일 cluster 로
        canonicalize 하기 위함). 없으면 새로 생성.
    layer_categories: 레이어→카테고리 매핑. None 이면 전체 entity 통과 (호환 모드).
    """
    g: dict[tuple[float, float], set[tuple[float, float]]] = defaultdict(set)
    edge_len: dict[tuple, float] = {}
    idx = node_index if node_index is not None else _NodeIndex()
    min_sq = min_edge_mm * min_edge_mm

    # 배관 geometry 가 될 수 없는 카테고리 — fallback 시에도 항상 제외.
    _NON_PIPE_CATS = {"HEAD", "TEXT", "ALARM", "ARCH", "EXCLUDE"}
    _pred_mode = "strict"   # "strict" → PIPE 만 / "broad" → 비-배관 카테고리만 제외

    def is_pipe(layer: str) -> bool:
        if layer_categories is None:
            return True
        cat = layer_categories.get(layer, "OTHER")
        if _pred_mode == "broad":
            return cat not in _NON_PIPE_CATS
        return cat == "PIPE"

    def add_edge(ax: float, ay: float, bx: float, by: float, length: float | None = None) -> None:
        a = idx.canonical(ax, ay)
        b = idx.canonical(bx, by)
        if a == b:
            return
        if length is None:
            length = math.hypot(b[0] - a[0], b[1] - a[1])
        if length < min_edge_mm:
            return
        g[a].add(b); g[b].add(a)
        key = (min(a, b), max(a, b))
        # 같은 노드 쌍에 더 짧은 edge_len 이 이미 있으면 덮어쓰지 않음 (실제 최단)
        prev = edge_len.get(key)
        if prev is None or length < prev:
            edge_len[key] = length

    def _consume() -> None:
        for en in pipe_entities:
            if not is_pipe(en.get("l", "")):
                continue
            et = en["t"]
            if et == "L":
                x1, y1, x2, y2 = en["p"]
                # 짧은 edge 사전 컷 (epsilon-cluster 전 raw 거리)
                if (x2 - x1) ** 2 + (y2 - y1) ** 2 < min_sq:
                    continue
                add_edge(x1, y1, x2, y2)
            elif et == "PL":
                pts = en["p"]
                if len(pts) < 2:
                    continue
                # closed polygon 감지 → 배관 아님
                first = pts[0]; last = pts[-1]
                if len(pts) >= 3 and math.hypot(first[0] - last[0], first[1] - last[1]) <= CLOSED_PL_TOL_MM:
                    continue
                for p0, p1 in zip(pts, pts[1:]):
                    if (p1[0] - p0[0]) ** 2 + (p1[1] - p0[1]) ** 2 < min_sq:
                        continue
                    add_edge(p0[0], p0[1], p1[0], p1[1])
            elif et == "A":
                # ARC — 양 끝점을 graph edge 로. 길이는 호 길이 (chord 아님).
                cx, cy = en["c"]
                r = float(en.get("r", 0.0) or 0.0)
                if r <= 0.0:
                    continue
                sa, ea = en.get("a", [0.0, 0.0])
                sa_r = math.radians(sa); ea_r = math.radians(ea)
                ax = cx + r * math.cos(sa_r); ay = cy + r * math.sin(sa_r)
                bx = cx + r * math.cos(ea_r); by = cy + r * math.sin(ea_r)
                # 호 sweep 각도 정규화 (0~360)
                sweep = ea - sa
                while sweep < 0: sweep += 360.0
                while sweep >= 360.0: sweep -= 360.0
                arc_len = r * math.radians(sweep)
                if arc_len < min_edge_mm:
                    continue
                add_edge(ax, ay, bx, by, length=arc_len)

    _consume()
    # PIPE-strict 가 빈 그래프를 내면 (배관이 '0' 등 OTHER 레이어에 작도된 도면)
    # 비-배관 카테고리만 제외하고 재시도. layer_categories=None 이면 두 모드가 동일해
    # 재시도해도 결과 불변이므로 strict 결과를 그대로 둔다.
    if not edge_len and layer_categories is not None:
        _pred_mode = "broad"
        _consume()
    return g, edge_len


# ────────────────────────────────────────────────────────────────────────────
# 평행 ladder collapse — 관경 두 줄 표현 → 중심선 1줄로 합성
# ────────────────────────────────────────────────────────────────────────────
# CAD 도면에서 배관을 두 평행 LINE 으로 그리는 관례 (관경 시각 표현).
# 그래프 빌드 시 두 줄이 모두 edge 로 남아 "ladder" (사다리) 모양 → 분기 수 부풀고
# 시각적으로 꼬임/겹침. 이 모듈은 4-cycle (u-v-w-x) 중 두 변이 평행하고 (rail)
# 나머지 두 변이 짧으면 (rung, 관 cap/cross-fitting) ladder 로 식별, midline 하나로 합성.







def _find_ladder_4cycles(
    graph: dict[tuple, set[tuple]],
    edge_len: dict[tuple, float],
    max_rung_mm: float,
    min_rail_ratio: float,
    cos_tol: float,
) -> list[tuple]:
    """4-cycle ladder 후보 검출.

    반환: (u, v, w, x, case) 리스트.
        case 'A': rails = (u,v) & (x,w), rungs = (v,w) & (x,u)
                  midline endpoints = mid(u,x), mid(v,w)
        case 'B': rails = (v,w) & (u,x), rungs = (u,v) & (w,x)
                  midline endpoints = mid(u,v), mid(w,x)
    """
    out: list[tuple] = []
    seen: set[frozenset] = set()

    def edge_length(a, b):
        key = (min(a, b), max(a, b))
        return edge_len.get(key, math.hypot(a[0] - b[0], a[1] - b[1]))

    for u in list(graph.keys()):
        nbs = list(graph.get(u, ()))
        n = len(nbs)
        if n < 2:
            continue
        for i in range(n):
            v = nbs[i]
            v_nb = graph.get(v, set()) - {u}
            if not v_nb:
                continue
            for j in range(i + 1, n):
                x = nbs[j]
                x_nb = graph.get(x, set()) - {u}
                if not x_nb:
                    continue
                common = v_nb & x_nb
                for w in common:
                    if w == u:
                        continue
                    cyc = frozenset((u, v, w, x))
                    if len(cyc) < 4 or cyc in seen:
                        continue
                    seen.add(cyc)

                    l_uv = edge_length(u, v)
                    l_vw = edge_length(v, w)
                    l_wx = edge_length(w, x)
                    l_xu = edge_length(x, u)

                    d_uv = _edge_dir(u, v)
                    d_xw = _edge_dir(x, w)
                    cos_A = abs(d_uv[0] * d_xw[0] + d_uv[1] * d_xw[1])

                    d_vw = _edge_dir(v, w)
                    d_ux = _edge_dir(u, x)
                    cos_B = abs(d_vw[0] * d_ux[0] + d_vw[1] * d_ux[1])

                    # Case A 와 B 모두 평가하고 rail/rung ratio 가 더 큰 쪽 선택.
                    # (cos 만 보고 case 결정하면 평행성은 더 좋지만 rung 길이 실패하는
                    #  쪽으로 빠질 수 있음 — T6 케이스)
                    chosen = None
                    best_ratio = 0.0
                    if cos_A >= cos_tol and l_vw <= max_rung_mm and l_xu <= max_rung_mm:
                        avg_rail = (l_uv + l_wx) / 2.0
                        avg_rung = (l_vw + l_xu) / 2.0
                        ratio = avg_rail / max(avg_rung, 1.0)
                        if ratio >= min_rail_ratio and ratio > best_ratio:
                            chosen = (u, v, w, x, "A")
                            best_ratio = ratio
                    if cos_B >= cos_tol and l_uv <= max_rung_mm and l_wx <= max_rung_mm:
                        avg_rail = (l_vw + l_xu) / 2.0
                        avg_rung = (l_uv + l_wx) / 2.0
                        ratio = avg_rail / max(avg_rung, 1.0)
                        if ratio >= min_rail_ratio and ratio > best_ratio:
                            chosen = (u, v, w, x, "B")
                            best_ratio = ratio
                    if chosen is not None:
                        out.append(chosen)
    return out


def _collapse_one_ladder(
    graph: dict[tuple, set[tuple]],
    edge_len: dict[tuple, float],
    u: tuple, v: tuple, w: tuple, x: tuple, case: str,
) -> tuple[tuple, tuple]:
    """4-cycle 의 네 노드를 제거하고 두 midpoint (m1, m2) + edge m1-m2 로 합성.

    cycle 노드의 외부 연결은 가까운 midpoint 로 redirect.
    return: (m1, m2) midpoint 좌표.
    """
    if case == "A":
        m1 = _midpoint(u, x)
        m2 = _midpoint(v, w)
        m1_src = (u, x)
        m2_src = (v, w)
    else:
        m1 = _midpoint(u, v)
        m2 = _midpoint(w, x)
        m1_src = (u, v)
        m2_src = (w, x)

    cycle_nodes = {u, v, w, x}
    # 외부 연결 수집 (cycle 내부 연결 제외)
    ext_m1: set[tuple] = set()
    ext_m2: set[tuple] = set()
    for s in m1_src:
        ext_m1.update(graph.get(s, set()) - cycle_nodes)
    for s in m2_src:
        ext_m2.update(graph.get(s, set()) - cycle_nodes)
    # 동일 노드가 양쪽에 — 매우 드물지만 m1 우선
    ext_m2 -= ext_m1

    # cycle 의 모든 edge_len 제거 (외부 연결 포함)
    for n in cycle_nodes:
        for nb in list(graph.get(n, ())):
            edge_len.pop((min(n, nb), max(n, nb)), None)
            graph[nb].discard(n)
        graph.pop(n, None)

    # m1, m2 노드 추가 + cycle 내부 midline edge
    graph[m1] = set(ext_m1) | {m2}
    graph[m2] = set(ext_m2) | {m1}
    edge_len[(min(m1, m2), max(m1, m2))] = math.hypot(m1[0] - m2[0], m1[1] - m2[1])
    for nb in ext_m1:
        graph.setdefault(nb, set()).add(m1)
        edge_len[(min(m1, nb), max(m1, nb))] = math.hypot(m1[0] - nb[0], m1[1] - nb[1])
    for nb in ext_m2:
        graph.setdefault(nb, set()).add(m2)
        edge_len[(min(m2, nb), max(m2, nb))] = math.hypot(m2[0] - nb[0], m2[1] - nb[1])
    return m1, m2


def collapse_parallel_ladders(
    graph: dict[tuple, set[tuple]],
    edge_len: dict[tuple, float],
    max_rung_mm: float = LADDER_MAX_RUNG_MM,
    min_rail_ratio: float = LADDER_MIN_RAIL_RATIO,
    cos_tol: float = LADDER_PARALLEL_COS,
    max_iter: int = LADDER_MAX_ITER,
) -> int:
    """모든 평행 ladder 를 안정 상태까지 반복 합성.

    한 4-cycle 합성이 인접 ladder 의 노드를 바꿀 수 있어 한 패스에서는 노드 중복
    사용을 피하고, 패스 사이에는 다시 검출. 새로 생성된 m1/m2 가 다음 패스의
    ladder 일부일 수 있어 max_iter 회 반복.

    return: 합성된 ladder 총 개수.
    """
    total = 0
    for _ in range(max_iter):
        candidates = _find_ladder_4cycles(graph, edge_len, max_rung_mm, min_rail_ratio, cos_tol)
        if not candidates:
            break
        applied = 0
        used: set[tuple] = set()
        for (u, v, w, x, case) in candidates:
            quad = {u, v, w, x}
            if quad & used:
                continue
            if not all(n in graph for n in quad):
                continue
            _collapse_one_ladder(graph, edge_len, u, v, w, x, case)
            used.update(quad)
            applied += 1
        total += applied
        if applied == 0:
            break
    return total


# ── 적응형 추정연결 허용치 — 도면 스케일(그래프 bbox 대각선)에 비례.
# 고정 mm 값은 소도면·대도면을 동시에 만족 못 한다: 같은 5m 가 대명동(대각선
# 44.6m)엔 11%(방 가로지르기 오접합)·B1F(783.6m)엔 0.6%(국소 이음매)라 정반대
# 성격이다. → tol = clamp(frac·diag, floor, ceil) 로 두 도면의 검증된 값을 동시
# 재현: weld 대명동≈2.0m/B1F=5.0m, bridge 상한 대명동≈3.1m/B1F=10m.
_WELD_TOL_FRAC = 0.045
_WELD_TOL_MIN = 500.0
_WELD_TOL_MAX = 5000.0
# cone 은 스케일 무관 — B1F 검증값 35° 유지. 보수화는 거리(tol)로만.
# (cone 을 25° 로 조이면 B1F 컴포넌트 20→28 로 조각남, 대명동 이득은 미미.)
_WELD_CONE_DEG = 35.0
_BRIDGE_TOP_FRAC = 0.07
_BRIDGE_TOP_MIN = 2000.0
_BRIDGE_TOP_MAX = 10000.0
_BRIDGE_BASE_TOLS = (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0)


def _graph_diag(graph: dict) -> float:
    """그래프 노드들의 bbox 대각선(mm). 빈 그래프면 0."""
    if not graph:
        return 0.0
    xs = [n[0] for n in graph]
    ys = [n[1] for n in graph]
    return math.hypot(max(xs) - min(xs), max(ys) - min(ys))


def _adaptive_weld_tol(diag: float) -> float:
    return min(_WELD_TOL_MAX, max(_WELD_TOL_MIN, _WELD_TOL_FRAC * diag))


def _adaptive_bridge_tols(diag: float) -> list[float]:
    """스케일 비례 bridge 사다리. base 스텝(고정 이음매 갭) + 상한(도면비례).
    B1F 처럼 diag 크면 base 그대로(200~10000), 대명동처럼 작으면 상한이 낮아져
    (예: 3.1m) 도면 가로지르는 장거리 강제연결을 원천 차단."""
    top = min(_BRIDGE_TOP_MAX, max(_BRIDGE_TOP_MIN, _BRIDGE_TOP_FRAC * diag))
    tols = [t for t in _BRIDGE_BASE_TOLS if t < top]
    tols.append(top)
    return tols


# ────────────────────────────────────────────────────────────────────────────
# Spanning Tree 강제 (가지식 트리 변환)
# ────────────────────────────────────────────────────────────────────────────
# 한국 NFTC 표준 SP 시스템은 기본 "가지식" — AV → 본관 → 가지관 → 헤드 트리 구조.
# 그래프에 cycle 이 남으면 (CAD 작도 실수, 미해결 ladder, 텍스트 box 오인 등)
# 시각적 겹침/꼬임 + hydraulic 계산 부정확 (cycle 안 분배 흐름).
# force_spanning_tree 는 AV-rooted Dijkstra SPT 로 강제 트리화.
# - 도달 가능한 노드: AV 까지 최단 경로 트리
# - 도달 불가능한 component: 각자 임의 root 의 SPT
# - 제거된 edge 들은 별도 set 으로 반환 → 시각화에서 cycle 자리 표시 가능


def _shortest_path_parents(
    graph: dict[tuple, set[tuple]],
    edge_len: dict[tuple, float],
    root: tuple,
    penalty_keys: set | None = None,
) -> dict[tuple, tuple]:
    """root 기준 min-weight 최단경로 부모 맵 {node: parent}. root 는 키에 없음.

    부모는 그 노드로 오는 tight-edge(최단경로 선행자 간선) 중 **가장 가벼운 것**으로
    고른다. 단순 Dijkstra 선행자는 relax 순서에 따라 무거운 shortcut 간선이 부모로
    남아 밸브 근처에 가짜 hub(조기분기)를 만드는데, tight-edge 최소 가중치 부모를
    쓰면 그런 가짜 분기가 사라져 주배관이 하나의 선으로 정돈된다
    (대명동 junction 235→116; B1F snaking 없음). 결과는 relax 순서와 무관.

    penalty_keys: 추정연결(weld/bridge 등) edge 키 집합. 주어지면 라우팅 비용을
        **사전식(추정edge 개수, 실길이)** 으로 정렬해 실배관을 우선 선택한다.
        추정 직선은 기하학적으로 짧아 순수 길이 기준으론 도면을 가로지르는
        지름길(wormhole)로 선택돼 경로가 꼬이는데, "추정 edge 를 하나라도 덜 밟는
        경로"를 항상 우선하면 실배관이 있으면 그쪽으로 돈다(추정 없이 도달 불가한
        구간만 추정 사용). **거리(수리계산)는 이후 실 edge_len 으로 재므로 물리
        거리는 보존.** 미지정 시 (추정 개수 항상 0) 순수 길이 기준. 가법 penalty
        (거대 상수) 대신 사전식을 쓰는 이유: 거대 상수는 누적 거리를 폭증시켜
        tight-edge 상대 허용치를 무너뜨려 트리가 끊긴다.
    """
    _pk = penalty_keys or set()

    def _w(u, v):
        w = edge_len.get((min(u, v), max(u, v)))
        return w if w is not None else math.hypot(u[0] - v[0], u[1] - v[1])

    def _pen(u, v):
        """이 edge 가 추정연결이면 1, 실배관이면 0 (사전식 라우팅 1차 키)."""
        return 1 if (_pk and (min(u, v), max(u, v)) in _pk) else 0

    _INF = (float("inf"), float("inf"))
    dist: dict[tuple, tuple] = {root: (0, 0.0)}
    pq: list[tuple] = [(0, 0.0, root)]
    while pq:
        pc, d, u = heapq.heappop(pq)
        if (pc, d) > dist[u]:
            continue
        for v in graph.get(u, ()):
            npc = pc + _pen(u, v)
            nd = d + _w(u, v)
            if (npc, nd) < dist.get(v, _INF):
                dist[v] = (npc, nd)
                heapq.heappush(pq, (npc, nd, v))
    # tight 판정: 추정 개수는 정확히 일치, 실길이는 상대 tol 이내(부동소수 오차).
    tol = 1e-6
    parents: dict[tuple, tuple] = {}
    for v, (pcv, dvl) in dist.items():
        if v == root:
            continue
        best_u = None
        best_w = float("inf")
        for u in graph.get(v, ()):
            du = dist.get(u)
            if du is None:
                continue
            pcu, dul = du
            w = _w(u, v)
            if pcu + _pen(u, v) == pcv and dul + w <= dvl + tol * (1.0 + abs(dvl)):
                if w < best_w or (w == best_w and (best_u is None or u < best_u)):
                    best_w = w
                    best_u = u
        if best_u is not None:
            parents[v] = best_u
    return parents


def water_load_audit(
    graph: dict[tuple, set[tuple]],
    edge_len: dict[tuple, float],
    source: tuple | None,
    heads: list,
    penalty_keys: set | None = None,
    max_attach_mm: float = HEAD_BRIDGE_MAX_MM,
) -> dict:
    """급수 감사 — 각 edge 가 몇 개의 승인 헤드에 물을 보내는지 계측(POC3 water_cleanup).

    source→각 헤드 최단경로(force_spanning_tree 와 동일한 사전식 비용)를 겹쳐 edge 별
    급수 부하를 센다. 부하 0 인 edge = 어떤 헤드에도 기여하지 않는 구간 —
    드레인·시험배관·표기 잔재 후보이자 추출 결함 신호다.

    집계 범위는 **source 도달 가능 컴포넌트로 한정**한다. 이 시점의 graph 는 도면
    전체(타 세대망·노이즈 조각 포함)를 담고 있어, 전역으로 재면 "우리 망과 무관한
    별개 컴포넌트"가 전부 부하 0 으로 잡혀 비율이 무의미해진다(대명동 실측 89%).

    **계측 전용** — graph/edge_len 을 수정하지 않는다. 부하 0 을 곧바로 삭제하면 안 되는
    이유는 최단경로만 겹치므로 루프의 대체 경로가 0 으로 잡히기 때문(영역 내 루프는
    보존 정책). 실제 가지치기는 실측 dead_ratio 를 본 뒤 별도 결정.

    SPT 이전(사이클 보유 상태)에 재는 것이 전제 — SPT 후에는 사이클 edge 가 이미
    제거돼 루프/드레인이 관측 대상에서 사라진다.
    """
    empty = {"dead_edge_count": 0, "dead_len_mm": 0.0, "dead_ratio": 0.0,
             "watered_edge_count": 0, "heads_routed": 0}
    if source is None or source not in graph:
        return empty
    parents = _shortest_path_parents(graph, edge_len, source, penalty_keys)
    load: dict[tuple, int] = defaultdict(int)
    routed = 0
    for h in heads:
        pos = h.pos if hasattr(h, "pos") else (float(h[0]), float(h[1]))
        near = _nearest_graph_node(graph, pos)
        if near is None or math.hypot(pos[0] - near[0], pos[1] - near[1]) > max_attach_mm:
            continue
        if near != source and near not in parents:
            continue  # source 와 다른 컴포넌트 — 미도달 헤드로 이미 보고됨
        routed += 1
        v = near
        seen = {v}
        while v in parents:
            u = parents[v]
            load[(min(u, v), max(u, v))] += 1
            if u in seen:
                break
            seen.add(u)
            v = u
    dead_n = 0
    dead_len = 0.0
    total_len = 0.0
    watered = 0
    reachable = set(parents)
    reachable.add(source)
    for key in {(min(u, v), max(u, v)) for u, nbs in graph.items() for v in nbs
                if u in reachable}:
        L = edge_len.get(key, 0.0)
        total_len += L
        if load.get(key):
            watered += 1
        else:
            dead_n += 1
            dead_len += L
    return {"dead_edge_count": dead_n, "dead_len_mm": dead_len,
            "dead_ratio": (dead_len / total_len) if total_len else 0.0,
            "watered_edge_count": watered, "heads_routed": routed}


def force_spanning_tree(
    graph: dict[tuple, set[tuple]],
    edge_len: dict[tuple, float],
    source: tuple | None = None,
    penalty_keys: set | None = None,
) -> tuple[set, set]:
    """그래프를 (각 component 마다) min-weight shortest-path spanning tree 로 변환.

    트리 간선 선택 규칙은 `_shortest_path_parents` 참조.

    Args:
        graph: 무방향 그래프 (in-place 수정됨 — cycle edge 제거)
        edge_len: edge 길이 dict (in-place 수정 — 제거된 edge 도 같이 pop)
        source: AV (또는 시작 노드). 이 노드가 속한 component 는 source 가 root.
                다른 component 는 임의 root (가장 작은 좌표 노드).
        penalty_keys: 추정연결 edge 키 집합 (사전식 라우팅). 미지정 시 순수 길이
                기준 = 기존 동작 불변(바이트 동일).

    Returns:
        (tree_edges, removed_edges) — 각각 (min, max) 키 set.
    """
    tree_edges: set[tuple] = set()
    for comp in _connected_components(graph):
        if not comp:
            continue
        if source is not None and source in comp:
            root = source
        else:
            root = min(comp, key=lambda p: (p[0], p[1]))  # deterministic
        for v, u in _shortest_path_parents(graph, edge_len, root, penalty_keys).items():
            tree_edges.add((min(v, u), max(v, u)))

    # 전체 edge 수집 → tree 외는 제거 대상
    all_edges: set[tuple] = set()
    for u, nbs in list(graph.items()):
        for v in nbs:
            all_edges.add((min(u, v), max(u, v)))
    removed_edges = all_edges - tree_edges

    # in-place 수정 — 트리 외 edge 제거
    for (a, b) in removed_edges:
        graph[a].discard(b)
        graph[b].discard(a)
        edge_len.pop((a, b), None)
    # 빈 인접 set 노드는 그대로 유지 (시각화에서 isolated 노드도 보여줘야 함)

    return tree_edges, removed_edges


def _restrict_to_branch_region(
    graph: dict[tuple, set[tuple]],
    edge_len: dict[tuple, float],
    source: tuple | None,
    branch_zones: "list[tuple[float, float, float, float]] | HeadRegion | None",
    penalty_keys: set | None = None,
) -> set:
    """분기영역(branch_zones) 지정 시 그래프를 in-place 로 제한한다.

    규칙(옵션 B — 표시만 보존, 계산은 트리):
      - 영역 **밖**: source → 영역 최단 단일 경로(corridor)만 남긴다. 나머지
        주배관 가지·루프는 제거 → 주배관→교차배관 구간은 배관 하나만.
      - 영역 **안**: 모든 edge 보존(루프 포함). 이후 SPT 가 계산용 트리를 만들되
        제거된 루프 edge 는 호출측이 실선(_graph_loop)으로 표시.

    branch_zones 없거나 도달 가능한 영역 노드 없으면 no-op(그래프 불변) →
    분기영역 미지정 도면(예: 대명동)은 전혀 영향 없음.

    Returns:
        region_nodes — 영역(분기영역 사각형 union) 안에 든 그래프 노드 set.
        빈 set 이면 아무것도 하지 않았다는 뜻(no-op).
    """
    if not branch_zones or source is None or source not in graph:
        return set()

    # W4: 영역 표현 통일 — rect list 는 HeadRegion.from_rects 로 승격.
    # 내부 의미론 불변: in_region 판정만 HeadRegion.contains 에 위임
    # (min/max 정규화 + 경계 포함 <= — 기존 판정과 비트동일).
    region = (branch_zones if isinstance(branch_zones, HeadRegion)
              else HeadRegion.from_rects(branch_zones))
    in_region = region.contains

    region_nodes = {n for n in graph if in_region(n)}
    if not region_nodes:
        return set()

    _pk = penalty_keys or set()

    def _w(u, v):
        w = edge_len.get((min(u, v), max(u, v)))
        return w if w is not None else math.hypot(u[0] - v[0], u[1] - v[1])

    def _pen(u, v):
        return 1 if (_pk and (min(u, v), max(u, v)) in _pk) else 0

    # source → 최단거리 영역 노드까지 사전식 Dijkstra + 경로 복원. 영역 노드에 처음
    # 도달하는 순간 종료 → 단일 corridor (영역 밖은 이 경로만 남는다).
    # 비용 (추정edge 개수, 실길이) — corridor 가 실배관을 우선(추정 최소).
    _INF = (float("inf"), float("inf"))
    dist = {source: (0, 0.0)}
    prev: dict[tuple, tuple] = {}
    pq = [(0, 0.0, source)]
    entry = source if source in region_nodes else None
    while pq and entry is None:
        pc, d, u = heapq.heappop(pq)
        if (pc, d) > dist[u]:
            continue
        if u in region_nodes:
            entry = u
            break
        for v in graph.get(u, ()):
            npc = pc + _pen(u, v)
            nd = d + _w(u, v)
            if (npc, nd) < dist.get(v, _INF):
                dist[v] = (npc, nd)
                prev[v] = u
                heapq.heappush(pq, (npc, nd, v))
    if entry is None:
        return set()  # 영역이 source 로부터 도달 불가 — 안전하게 no-op

    corridor_edges: set = set()
    cur = entry
    while cur in prev:
        p = prev[cur]
        corridor_edges.add((min(cur, p), max(cur, p)))
        cur = p

    # 보존 = 영역 내부 edge(양 끝 모두 영역) ∪ corridor. 나머지 제거.
    keep: set = set()
    all_edges: set = set()
    for u, nbs in graph.items():
        for v in nbs:
            key = (min(u, v), max(u, v))
            all_edges.add(key)
            if (u in region_nodes and v in region_nodes) or key in corridor_edges:
                keep.add(key)
    for (a, b) in (all_edges - keep):
        graph[a].discard(b)
        graph[b].discard(a)
        edge_len.pop((a, b), None)

    return region_nodes


def _find_head_candidates(pipe_entities: list[dict], layer_categories: dict[str, str],
                          region=None) -> list[HeadCandidate]:
    """자동 헤드 후보 — Stage 2 ``detect_heads`` 결과를 그대로 사용.

    과거엔 HEAD 레이어의 INSERT/CIRCLE 만 모으는 별도 단순 로직이었으나, 그러면
    사용자가 Stage 2 에서 보는 헤드 집합(detect_heads: HATCH 삼각형·layer-agnostic
    드라이팬던트·근접 클러스터링 포함)과 자동 fallback 경로가 어긋났다(예: 대명동
    32 vs 28). select_worst30_heads 가 manual_heads 없이 호출될 때도 동일 집합을
    쓰도록 detect_heads 로 위임한다. raw=pos (클러스터 대표 cue 좌표).
    region: anchored 영역 게이트 — detect_heads 로 그대로 전달(W1). None 이면 불변.
    """
    detections = detect_heads(pipe_entities, layer_categories, region=region)
    return [
        HeadCandidate(pos=_round_pt(d.pos[0], d.pos[1]), raw=d.pos,
                      block_name=d.block_name or f"({d.kind})", layer=d.layer)
        for d in detections
    ]


def _find_source(pipe_entities: list[dict], layer_categories: dict[str, str]) -> tuple[tuple[float, float] | None, str]:
    """알람밸브 자동 식별 — 5-tier fallback:
      1) block_name 에 ALARM_VALVE 키워드 포함된 INSERT (사전 기반)
      2) layer 이름이 ALARM_VALVE 키워드 포함된 INSERT (예: 'RISER' 레이어)
      3) '배관-SP 2차' 또는 'SP 2차' 레이어의 첫 INSERT (입상→알람→가지 source)
      4) '배관-SP 2차' 레이어의 LINE 의 endpoint 중 가지관 그래프와 가장 가까운 점
      5) None (호출자가 fallback 처리)
    """
    # 사전 import — 사용자가 sprinkler_remote30_extractor.py 에서 키워드 추가하면 자동 반영
    try:
        from sprinkler_remote30_extractor import DEFAULT_ALARM_VALVE_KEYWORDS as _AV_KW
        av_keywords = [k.upper() for k in _AV_KW]
    except ImportError:
        av_keywords = ["ALARM", "알람", "알람밸브", "RISER", "라이저",
                        "STAND-PIPE", "STANDPIPE", "STAND_PIPE"]

    def _matches_av(text: str) -> bool:
        up = (text or "").upper()
        return any(kw in up for kw in av_keywords)

    # tier 1: block_name 매칭
    for en in pipe_entities:
        if en["t"] != "I":
            continue
        if _matches_av(en.get("n") or ""):
            return _round_pt(en["p"][0], en["p"][1]), "alarm_block"
    # tier 2: layer 이름 매칭 (예: 'RISER' 레이어의 INSERT) — 새로 추가
    for en in pipe_entities:
        if en["t"] != "I":
            continue
        if _matches_av(en.get("l") or ""):
            return _round_pt(en["p"][0], en["p"][1]), "alarm_layer"
    # tier 3: 2차측 배관 레이어의 INSERT
    for en in pipe_entities:
        if en["t"] != "I":
            continue
        if "배관-SP 2차" in en["l"] or "SP 2차" in en["l"]:
            return _round_pt(en["p"][0], en["p"][1]), "secondary_layer_insert"
    # tier 4: 2차 배관 LINE 의 endpoint 들 수집
    secondary_endpoints: list[tuple[float, float]] = []
    for en in pipe_entities:
        if en["t"] == "L" and ("배관-SP 2차" in en["l"] or "SP 2차" in en["l"]):
            p = en["p"]
            secondary_endpoints.append(_round_pt(p[0], p[1]))
            secondary_endpoints.append(_round_pt(p[2], p[3]))
    if secondary_endpoints:
        from collections import Counter as _C
        ec = _C(secondary_endpoints)
        return ec.most_common(1)[0][0], "secondary_layer_line"
    return None, "auto_junction"










def _bridge_components(
    graph: dict,
    edge_len: dict,
    max_bridge_mm: float = 500.0,
    bridge_edges_out: set | None = None,
) -> int:
    """끊어진 component 들을 가장 가까운 endpoint 쌍 연결 — 50cm 이내만.

    bridge_edges_out: 주어지면 추가된 bridge edge 의 (min,max) 키를 누적.
        호출자가 "실제 배관"과 "알고리즘이 추정한 연결"을 구분 렌더할 수 있음.
    """
    # 한 번의 pass 만 돌면 stale-main chain 문제 발생: A→main 으로 병합되며
    # main 이 커지면 그 다음에야 tol 안에 드는 comp(B) 가 같은 pass 에선 이미
    # 평가·skip 된 뒤라 연결되지 않음 (예: LH306 입상 — 알람밸브 stub 가 중간
    # 입상 comp 가 main 에 붙은 직후 1.7m 거리로 좁혀지지만 그 pass 는 끝나버림).
    # → comp 를 재계산하며 더 이상 병합이 없을 때까지 반복 (single-linkage).
    total = 0
    while True:
        comps = _connected_components(graph)
        if len(comps) <= 1:
            break
        # main = 가장 큰 component
        main = max(comps, key=len)
        others = [c for c in comps if c is not main]
        bridges = 0
        # 공간 격자 인덱스 — cell = max_bridge_mm. tol 이내 후보는 반드시
        # 자기 셀의 3x3 이웃 안에 있으므로, 브루트포스 O(|comp|*|main|) 를
        # 근사 O(|comp|) 로 줄인다 (조각난 대형 도면에서 분→초).
        # 결과는 브루트포스와 "바이트 동일" — 등거리 tie 는 원본과 같이
        #   ① 같은 u 안에서는 main 반복순서(main_index) 가 앞선 v,
        #   ② u 간에는 먼저 최소를 달성한 u (strict < 비교) 를 택해 재현.
        inv = 1.0 / max_bridge_mm
        grid: dict[tuple[int, int], list] = defaultdict(list)
        main_index: dict = {}
        for i, v in enumerate(main):
            grid[(int(math.floor(v[0] * inv)), int(math.floor(v[1] * inv)))].append(v)
            main_index[v] = i
        tol2 = max_bridge_mm * max_bridge_mm
        for comp in others:
            # comp 의 각 노드에서 tol 이내 가장 가까운 main 노드 (제곱거리 비교).
            best = None
            bestd2 = float("inf")
            for u in comp:
                ux = u[0]; uy = u[1]
                cgx = int(math.floor(ux * inv)); cgy = int(math.floor(uy * inv))
                u_best_v = None; u_best_d2 = float("inf"); u_best_idx = -1
                for dgx in (-1, 0, 1):
                    for dgy in (-1, 0, 1):
                        for v in grid.get((cgx + dgx, cgy + dgy), ()):
                            dx = ux - v[0]; dy = uy - v[1]
                            d2 = dx * dx + dy * dy
                            if d2 < u_best_d2:
                                u_best_d2 = d2; u_best_v = v; u_best_idx = main_index[v]
                            elif d2 == u_best_d2 and main_index[v] < u_best_idx:
                                u_best_v = v; u_best_idx = main_index[v]
                if u_best_v is not None and u_best_d2 < bestd2:  # strict → 앞선 u 우선
                    bestd2 = u_best_d2
                    best = (u, u_best_v)
            if best is not None and bestd2 <= tol2:
                u, v = best
                bestd = math.hypot(u[0] - v[0], u[1] - v[1])  # 최종 1회 — 저장값 동일 보존
                graph[u].add(v); graph[v].add(u)
                key = (min(u, v), max(u, v))
                edge_len[key] = bestd
                if bridge_edges_out is not None:
                    bridge_edges_out.add(key)
                bridges += 1
        total += bridges
        if bridges == 0:
            break
    return total


def _split_tee_branches(
    graph: dict,
    edge_len: dict,
    max_gap_mm: float = TEE_SPLIT_MAX_MM,
    min_edge_mm: float = MIN_PIPE_EDGE_MM,
    splits_out: list | None = None,
) -> int:
    """T분기 복원 — 느슨한 끝점이 다른 배관 *중간* 에 닿는 곳에서 그 배관을 쪼갠다.

    _weld_dangling_endpoints 는 노드↔노드만 잇고 edge 를 쪼갤 수 없다. 실제 도면의
    가지관은 주배관 끝점이 아니라 중간에서 갈라지므로(T분기), weld 는 ① 엉뚱한
    조각에 붙이거나 ② 같은 배관이라도 그 *끝점* 에 붙여 배관 길이를 부풀린다
    (대명동 서측 세대 실측: 후보 47건 중 오접합 32건 / 제 edge 에 붙은 12건도
    중앙값 1.1m·합계 10.7m 의 가공 배관을 만든다).

    끝점 u 에서 edge (a,b) 내부로 내린 수선발이 max_gap_mm 안이면 (a,b) 를 지우고
    (a,u),(u,b) 로 대체한다. 분기점은 수선발이 아니라 u 자신 — 둘의 거리가
    max_gap_mm(≤ _NodeIndex epsilon) 안이라 이미 같은 노드로 취급되는 위치이고,
    새 좌표를 만들지 않아 raw DXF 좌표 보존 원칙이 지켜진다.

    안전장치:
      - degree-1 노드만 소스 (배관 중간점끼리 엮이는 사고 방지)
      - u 가 그 edge 의 끝점이거나 이미 인접이면 제외 (weld 영역)
      - 쪼갠 두 조각 모두 min_edge_mm 이상 — 최소길이 미만 edge 생성 금지
      - 새 edge 는 실배관이므로 추정연결(penalty) 아님
    splits_out: 주어지면 [{"p", "edge", "gap_mm"}] 기록 (audit 근거용).
    반환: 쪼갠 edge 수.
    """
    if max_gap_mm <= 0:
        return 0
    cell = max(500.0, max_gap_mm * 10.0)
    inv = 1.0 / cell
    grid: dict[tuple[int, int], list] = defaultdict(list)

    def _register(key: tuple) -> None:
        # 셀 절반 간격 표본 등록 — 3x3 조회와 합쳐 max_gap 안의 구간을 놓치지 않는다.
        (ax, ay), (bx, by) = key
        n = int(math.hypot(bx - ax, by - ay) * 2.0 * inv) + 1
        for i in range(n + 1):
            t = i / n
            grid[(int(math.floor((ax + t * (bx - ax)) * inv)),
                  int(math.floor((ay + t * (by - ay)) * inv)))].append(key)

    for key in edge_len:
        _register(key)

    splits = 0
    for u in sorted(n for n, nb in graph.items() if len(nb) == 1):
        nb = graph.get(u)
        if not nb or len(nb) != 1:
            continue
        ux, uy = u[0], u[1]
        cgx = int(math.floor(ux * inv)); cgy = int(math.floor(uy * inv))
        cands: list = []
        for dgx in (-1, 0, 1):
            for dgy in (-1, 0, 1):
                cands.extend(grid.get((cgx + dgx, cgy + dgy), ()))
        best = None; best_d = float("inf")
        for key in cands:
            a, b = key
            if key not in edge_len or b not in graph.get(a, ()):
                continue  # 앞선 split 으로 사라진 edge
            if u == a or u == b or a in nb or b in nb:
                continue
            abx = b[0] - a[0]; aby = b[1] - a[1]
            L2 = abx * abx + aby * aby
            if L2 < 1e-9:
                continue
            t = ((ux - a[0]) * abx + (uy - a[1]) * aby) / L2
            if not (0.0 < t < 1.0):
                continue  # 수선발이 edge 밖 — 끝점 접속(weld 영역)
            d = math.hypot(ux - (a[0] + t * abx), uy - (a[1] + t * aby))
            if d > max_gap_mm:
                continue
            if best is not None and (d, key) >= (best_d, best):
                continue  # 동점은 edge 키 순으로 결정적 해소
            if (math.hypot(ux - a[0], uy - a[1]) < min_edge_mm
                    or math.hypot(ux - b[0], uy - b[1]) < min_edge_mm):
                continue
            best = key; best_d = d
        if best is None:
            continue
        a, b = best
        graph[a].discard(b); graph[b].discard(a)
        del edge_len[best]
        for v in (a, b):
            graph[u].add(v); graph[v].add(u)
            k2 = (min(u, v), max(u, v))
            L = math.hypot(ux - v[0], uy - v[1])
            prev = edge_len.get(k2)
            if prev is None or L < prev:
                edge_len[k2] = L
            _register(k2)
        if splits_out is not None:
            splits_out.append({"p": [ux, uy], "gap_mm": best_d,
                               "edge": [[a[0], a[1]], [b[0], b[1]]]})
        splits += 1
    return splits


def _weld_dangling_endpoints(
    graph: dict,
    edge_len: dict,
    weld_tol: float = 5000.0,
    weld_cone_deg: float = 35.0,
    weld_edges_out: set | None = None,
) -> int:
    """끊긴 배관 방향-인지 복원 — degree-1(느슨한 끝) 노드를 그 배관 축을 따라
    전방 각도콘 안에서 이어지는 조각에 직접 연결(collinear continuation).

    _bridge_components 는 "가장 큰 덩어리에 가까운 조각부터" 잇는 main-centric
    single-linkage 라, 이음매마다 작은 갭으로 끊긴 현장조사 도면에서 파편을
    뱀처럼 이어붙여 밸브→헤드 경로가 크게 돌아간다(B1F 실측 8~40배). 단순
    최근접 용접/브리지는 방향을 몰라 옆 가지관에 붙거나 지그재그를 만든다.
    끝점의 outgoing 방향(u-w)을 따라 전방 콘 안의 노드만 후보로 삼아 먼저
    용접하면 끊긴 직선 배관이 원래 축대로 복원돼 경로가 곧아진다
    (B1F 밸브→헤드 8.85배 → 1.47배).

    안전장치:
      - degree-1 노드만 소스 (배관 중간점끼리 엮이는 사고 방지)
      - 전방 각도콘(기본 ±35°)만 후보 — 옆으로 나란한 평행 가지관(축에서
        ~90° 벗어남)에는 절대 붙지 않는다.
      - 거리(weld_tol)는 프로덕션 호출부에서 도면 스케일 비례 적응형으로 넘긴다
        (_adaptive_weld_tol): 대명동≈2m·B1F=5m. 소도면의 도면 가로지르기 장거리
        용접을 차단하면서 대도면의 국소 이음매 복원은 유지. (여기 기본값 5m 는
        직접 호출용 fallback.)
      - 정렬(cosang) 우선 + 근접 가산 점수로 가장 곧게 잇는 조각 선택
      - 용접이 새 degree-1 을 만들 수 있어 수렴할 때까지 반복(상한 8 pass)
    weld_edges_out: 주어지면 추가된 weld edge (min,max) 키 누적 (추정연결 렌더 구분용).
    반환: 추가한 weld edge 총수.
    """
    cos_min = math.cos(math.radians(weld_cone_deg))
    inv = 1.0 / weld_tol
    tol2 = weld_tol * weld_tol
    welds = 0
    for _ in range(8):
        dangling = [n for n, nb in graph.items() if len(nb) == 1]
        if not dangling:
            break
        grid: dict[tuple[int, int], list] = defaultdict(list)
        for n in graph:
            grid[(int(math.floor(n[0] * inv)), int(math.floor(n[1] * inv)))].append(n)
        pass_welds = 0
        for u in dangling:
            nb = graph.get(u)
            if not nb or len(nb) != 1:
                continue  # 이번 pass 중 이미 용접돼 degree 변함
            w = next(iter(nb))
            ux, uy = u[0], u[1]
            dx0 = ux - w[0]; dy0 = uy - w[1]
            L0 = math.hypot(dx0, dy0)
            if L0 < 1e-6:
                continue
            dx0 /= L0; dy0 /= L0  # outward unit direction of this pipe
            cgx = int(math.floor(ux * inv)); cgy = int(math.floor(uy * inv))
            best = None; best_score = -1e18; best_d2 = float("inf")
            for dgx in (-1, 0, 1):
                for dgy in (-1, 0, 1):
                    for v in grid.get((cgx + dgx, cgy + dgy), ()):
                        if v is u or v in nb:
                            continue
                        vx = v[0] - ux; vy = v[1] - uy
                        d2 = vx * vx + vy * vy
                        if d2 > tol2 or d2 < 1e-9:
                            continue
                        dd = math.sqrt(d2)
                        cosang = (vx * dx0 + vy * dy0) / dd
                        if cosang < cos_min:
                            continue  # 전방 콘 밖 (옆/뒤 조각 배제)
                        # 정렬 우선 + 근접 가산; 동점은 좌표순(v)으로 결정적 해소
                        score = cosang - (dd * inv) * 0.15
                        if score > best_score or (
                                score == best_score and (best is None or v < best)):
                            best_score = score; best = v; best_d2 = d2
            if best is not None:
                d = math.sqrt(best_d2)
                graph[u].add(best); graph[best].add(u)
                key = (min(u, best), max(u, best))
                edge_len[key] = d
                if weld_edges_out is not None:
                    weld_edges_out.add(key)
                welds += 1; pass_welds += 1
        if pass_welds == 0:
            break
    return welds


def select_worst30_heads(
    pipe_entities: list[dict],
    layer_categories: dict[str, str],
    k: int = 30,
    manual_source: tuple[float, float] | None = None,
    manual_heads: list[tuple[float, float]] | None = None,
    zones: list[tuple[float, float, float, float]] | None = None,
    branch_zones: list[tuple[float, float, float, float]] | None = None,
    progress_cb=None,
) -> SelectionResult:
    """가장 불리한 K 헤드 + 경로 선정.

    manual_heads: 명시되면 자동 검출 대신 이 리스트 사용 (사용자 편집 후)
    zones: [(x1,y1,x2,y2), ...] 영역 union. 비어있지 않으면 그 안의 헤드만 후보로.
    branch_zones: [(x1,y1,x2,y2), ...] 분기영역 union. 지정되면 영역 밖은
        source→영역 단일 최단경로만, 영역 안은 루프 보존(계산은 트리). 미지정 시 불변.
    progress_cb: 있으면 진행상황 콜백 progress_cb(fraction:0~1, label:str).
        서브단계 경계에서만 호출 — 산출 데이터는 건드리지 않아 결과 불변.
    """
    _pcb = progress_cb if callable(progress_cb) else (lambda f, m: None)
    _pcb(0.0, "배관망 그래프 구성 중")
    graph, edge_len = _build_graph(pipe_entities, layer_categories=layer_categories)
    # 평행 ladder collapse — Stage 3 시각화와 같은 토폴로지로 정렬.
    collapse_parallel_ladders(graph, edge_len)
    # 끊긴 배관 국소 복원 — main-centric 브리지 이전에 끝점끼리 국소 용접해
    # 원래 토폴로지 복원 (파편화 도면에서 뱀 경로 방지). 큰 갭은 이후 브리지가 처리.
    # 추정연결 허용치는 도면 스케일 비례(적응형) — 소도면 과잉연결·대도면 조각남 동시 방지.
    _diag = _graph_diag(graph)
    # 추정연결 edge 키 누적 — 이후 SPT/corridor 라우팅에서 penalty 부여(실배관 우선).
    _penalty_keys: set = set()
    _weld_dangling_endpoints(graph, edge_len,
                             weld_tol=_adaptive_weld_tol(_diag),
                             weld_cone_deg=_WELD_CONE_DEG,
                             weld_edges_out=_penalty_keys)
    _pcb(0.05, "평행 배관 정리 완료")
    # 짧은 거리부터 단계적으로 brigde — 가까운 endpoint 우선 + 점점 멀리.
    # 상한은 도면비례(대명동≈3.1m / B1F=10m) — 측지좌표 대도면은 5m/10m 밴드 유지,
    # 소도면은 상한이 낮아져 도면 가로지르는 장거리 강제연결 차단.
    _bridge_tols = _adaptive_bridge_tols(_diag)
    for _bi, tol in enumerate(_bridge_tols, 1):
        _bridge_components(graph, edge_len, max_bridge_mm=tol,
                           bridge_edges_out=_penalty_keys)
        _pcb(0.05 + 0.65 * _bi / len(_bridge_tols),
             f"조각난 배관 연결 {_bi}/{len(_bridge_tols)} (≤{int(tol)}mm)")
    # SPT 는 source 확정 후로 미룸 — source 루트 shortest-path tree 여야
    # 트리 경로 = 실제 최단경로가 되어, 용접/브리지가 추가한 edge 가 경로를
    # 늘리지 않는다(클린 도면 불변·파편 도면만 개선).
    _pcb(0.72, "알람밸브 식별 중")
    if manual_heads is not None:
        # 사용자가 편집한 헤드 목록 사용
        heads = [HeadCandidate(pos=_round_pt(x, y), raw=(x, y), block_name="(user)", layer="_user")
                 for x, y in manual_heads]
    else:
        heads = _find_head_candidates(pipe_entities, layer_categories)
    # zone 필터 — union 안에 들어오는 헤드만
    if zones:
        def in_any_zone(x: float, y: float) -> bool:
            for (zx1, zy1, zx2, zy2) in zones:
                lo_x, hi_x = (zx1, zx2) if zx1 <= zx2 else (zx2, zx1)
                lo_y, hi_y = (zy1, zy2) if zy1 <= zy2 else (zy2, zy1)
                if lo_x <= x <= hi_x and lo_y <= y <= hi_y:
                    return True
            return False
        heads = [h for h in heads if in_any_zone(h.pos[0], h.pos[1])]
    # K 도 적응형 — 헤드 수 부족하면 있는 만큼
    if len(heads) < k:
        k = len(heads)
    if manual_source is not None:
        src_raw = _round_pt(manual_source[0], manual_source[1])
        src_kind = "manual"
    else:
        src_raw, src_kind = _find_source(pipe_entities, layer_categories)

    src_nearest = _nearest_graph_node(graph, src_raw) if src_raw else None
    src_bridge_dist_mm = 0.0
    src_fallback = False
    if src_nearest is None:
        # fallback — 그래프 자체가 빈 경우 / src_raw 없음
        if graph:
            src = max(graph, key=lambda n: len(graph[n]))
            src_kind = "highest_degree"
        else:
            return SelectionResult(None, "none", [], [], [], [], {})
    else:
        d_src = math.hypot(src_raw[0] - src_nearest[0], src_raw[1] - src_nearest[1])
        src_bridge_dist_mm = d_src
        if d_src <= 1e-3:
            # 사용자 좌표가 정확히 그래프 노드 위 — nearest 그대로 사용
            src = src_nearest
        elif d_src <= SOURCE_BRIDGE_MAX_MM:
            # 한도 이내 — src_raw 를 그래프 노드로 추가하고 nearest 와 edge 로 연결
            graph.setdefault(src_raw, set()).add(src_nearest)
            graph[src_nearest].add(src_raw)
            _src_key = (min(src_raw, src_nearest), max(src_raw, src_nearest))
            edge_len[_src_key] = d_src
            _penalty_keys.add(_src_key)  # source 접속선도 추정연결 — 라우팅 penalty
            src = src_raw
        else:
            # 한도 초과 — nearest 로 fallback 하고 경고 플래그
            src = src_nearest
            src_fallback = True
            src_kind = src_kind + ":fallback_far"

    # 분기영역 지정 시 — 영역 밖을 source→영역 단일 corridor 로 제한(주배관 하나).
    # SPT 이전에 적용해야 트리가 단일 corridor 기준으로 정돈된다. 미지정 시 no-op.
    _restrict_to_branch_region(graph, edge_len, src, branch_zones,
                               penalty_keys=_penalty_keys)

    # 가지식 트리 강제 (SPT) — source 루트 shortest-path tree. source 를 확정한
    # 뒤 그 노드에 루팅해야 트리 경로 = 실제 최단경로. 헤드는 이후 leaf 로 부착
    # (degree-1 drop line 이라 cycle 없음). Stage 3 시각화와 동일 토폴로지 유지.
    # penalty_keys: 추정연결은 라우팅 비용에 penalty — 트리가 실배관 우선 선택.
    force_spanning_tree(graph, edge_len, source=src, penalty_keys=_penalty_keys)

    return _finalize_selection(graph, edge_len, src, src_kind, heads, k, _pcb,
                               src_bridge_dist_mm, src_fallback)


def _finalize_selection(
    graph: dict,
    edge_len: dict,
    src: tuple[float, float],
    src_kind: str,
    heads: list,
    k: int,
    _pcb,
    src_bridge_dist_mm: float,
    src_fallback: bool,
    head_drop_out: set | None = None,
) -> SelectionResult:
    """SPT 이후 공통 후반부 — 헤드 스냅→거리정렬→top-K→subgraph→collinear 병합.

    select_worst30_heads 에서 순수 코드 이동(동작 불변). anchored 경로
    (select_worst30_heads_anchored)와 공유하기 위해 분리.
    head_drop_out: 주어지면 헤드 스냅으로 추가된 head-drop edge (min,max) 키 누적
        (→W7 audit. 기본 None=기존 동작 불변).
    """
    # 헤드 최근접-노드 스냅 — 헤드 수천 개 × O(|graph|) 전수스캔이면 대형 도면에서
    # 수십 초~분. 격자 인덱스(cell = HEAD_BRIDGE_MAX_MM)로 근사 O(1) 질의로 대체.
    # 헤드 삽입 전 pipe-network 노드로만 그리드를 만든다(헤드끼리 스냅 방지).
    # node_order = graph 반복순서. 원본 _nearest_graph_node 는 `for n in graph`
    # strict < 스캔이라 등거리 시 graph 순서상 앞 노드를 택함 — 이를 재현해 동점 해소.
    _hg_inv = 1.0 / HEAD_BRIDGE_MAX_MM
    _hgrid: dict[tuple[int, int], list] = defaultdict(list)
    _hnode_order: dict = {}
    for _i, _n in enumerate(graph):
        _hgrid[(int(math.floor(_n[0] * _hg_inv)), int(math.floor(_n[1] * _hg_inv)))].append(_n)
        _hnode_order[_n] = _i
    _hnext_order = len(_hnode_order)  # 다음 삽입 노드의 graph 순서 인덱스

    def _hgrid_add(n):
        """헤드가 graph 에 삽입되면 그리드/순서에도 반영 — 원본은 graph 가
        커지며 뒤 헤드가 앞서 추가된 헤드 노드에 스냅될 수 있다(성장 그래프)."""
        nonlocal _hnext_order
        if n in _hnode_order:
            return
        _hgrid[(int(math.floor(n[0] * _hg_inv)), int(math.floor(n[1] * _hg_inv)))].append(n)
        _hnode_order[n] = _hnext_order
        _hnext_order += 1

    def _nearest_via_grid(pt, max_dist=None):
        """격자 이웃 링을 넓혀가며 pt 최근접 노드. max_dist 지정 시 그 이내만.
        등거리는 _hnode_order(=graph 순서) 앞선 노드 — 원본과 동일."""
        px, py = pt[0], pt[1]
        cgx = int(math.floor(px * _hg_inv)); cgy = int(math.floor(py * _hg_inv))
        best = None; bestd2 = float("inf"); best_ord = float("inf")
        r = 0
        while True:
            for gx in range(cgx - r, cgx + r + 1):
                for gy in range(cgy - r, cgy + r + 1):
                    if max(abs(gx - cgx), abs(gy - cgy)) != r:
                        continue  # 껍질(ring)만
                    for v in _hgrid.get((gx, gy), ()):
                        dx = px - v[0]; dy = py - v[1]; d2 = dx * dx + dy * dy
                        if d2 < bestd2:
                            bestd2 = d2; best = v; best_ord = _hnode_order[v]
                        elif d2 == bestd2:
                            o = _hnode_order[v]
                            if o < best_ord:
                                best = v; best_ord = o
            reach = r * HEAD_BRIDGE_MAX_MM  # 이 링까지 커버한 최소 반경
            if max_dist is not None and reach > max_dist:
                break  # 더 넓혀도 max_dist 밖 → 없음
            # strict > : 등거리(경계) 노드가 다음 링에 있을 수 있어 한 링 더 스캔 보장.
            if best is not None and reach > math.sqrt(bestd2):
                break
            r += 1
            if r > 4096:
                break  # 안전장치(사실상 도달 안 함)
        if max_dist is not None and best is not None and bestd2 > max_dist * max_dist:
            return None
        return best

    # 헤드 좌표 → 가장 가까운 그래프 노드로 강제 연결 (HEAD_BRIDGE_MAX_MM 이내)
    for h in heads:
        if h.pos in graph:
            continue  # 원본 _nearest_graph_node: pt in graph → pt(d=0) → 스킵
        nearest = _nearest_via_grid(h.pos, max_dist=HEAD_BRIDGE_MAX_MM)
        if nearest is None:
            continue
        d = math.hypot(h.pos[0] - nearest[0], h.pos[1] - nearest[1])
        if d > 1e-3 and d <= HEAD_BRIDGE_MAX_MM:
            graph.setdefault(h.pos, set()).add(nearest)
            graph[nearest].add(h.pos)
            edge_len[(min(h.pos, nearest), max(h.pos, nearest))] = d
            if head_drop_out is not None:
                head_drop_out.add((min(h.pos, nearest), max(h.pos, nearest)))
            _hgrid_add(h.pos)  # 성장 그래프 — 뒤 헤드가 이 노드에 스냅 가능

    _pcb(0.76, "알람밸브→전체 헤드 거리 계산 중")
    dist_map = _dijkstra_from(graph, edge_len, src)

    # head 후보들을 그래프 노드로 스냅 후 거리 정렬 — 도달 불가도 가능한 한 포함
    head_with_d: list[tuple[HeadCandidate, tuple[float, float], float]] = []
    for h in heads:
        node = h.pos if h.pos in graph else _nearest_via_grid(h.pos, max_dist=HEAD_BRIDGE_MAX_MM)
        if node is None:
            continue
        d = dist_map.get(node, float("inf"))
        if math.isfinite(d):
            head_with_d.append((h, node, d))
    head_with_d.sort(key=lambda x: -x[2])  # 멀리 있는 순
    top_k = head_with_d[:k]

    selected_heads = [h for h, _, _ in top_k]
    distances = [d for _, _, d in top_k]

    # subgraph 추출 — top-K 헤드 각각의 src→head 최단경로의 합집합
    sub_edges_seen: set[tuple[tuple[float, float], tuple[float, float]]] = set()
    sub_edges: list[tuple[tuple[float, float], tuple[float, float], float]] = []
    sub_nodes: set[tuple[float, float]] = {src}
    _n_top = len(top_k)
    for _si, (_, head_node, _) in enumerate(top_k, 1):
        path = _shortest_path(graph, edge_len, src, head_node)
        for a, b in zip(path, path[1:]):
            key = (min(a, b), max(a, b))
            if key in sub_edges_seen:
                continue
            sub_edges_seen.add(key)
            sub_edges.append((a, b, edge_len.get(key, math.hypot(b[0] - a[0], b[1] - a[1]))))
            sub_nodes.add(a); sub_nodes.add(b)
        if _si % 3 == 0 or _si == _n_top:
            _pcb(0.78 + 0.20 * _si / max(1, _n_top),
                 f"가장 불리한 경로 추적 {_si}/{_n_top}")

    # ====== Collinear merge — 직선상 degree-2 노드 제거 ======
    # source / heads / 차수≥3 노드는 절대 보존, 직선상 degree-2 노드만 흡수
    head_positions = {h.pos for h in selected_heads}
    keep_nodes = {src} | head_positions

    sub_adj: dict[tuple[float, float], list[tuple[float, float]]] = defaultdict(list)
    sub_edge_len: dict[tuple, float] = {}
    for a, b, L in sub_edges:
        sub_adj[a].append(b); sub_adj[b].append(a)
        sub_edge_len[(min(a, b), max(a, b))] = L

    def _angle(p, q):
        return math.atan2(q[1] - p[1], q[0] - p[0])

    # 1) 직선 흡수 — angle 차이 ≤ COLLINEAR_TOL, 자동 흡수
    # 2) 짧은 segment 흡수 — degree-2 + 두 segment 모두 SHORT_SEG_MM 이내, angle 차이 ≤ ELBOW_MERGE_TOL, 자동 흡수
    # 그 외 elbow 는 edge_elbows 로 기록 → fitting 으로 별도 보존
    COLLINEAR_TOL_DEG = 12.0
    ELBOW_MERGE_TOL_DEG = 95.0
    SHORT_SEG_MM = 500.0

    edge_elbows: dict[tuple, list[tuple[tuple[float, float], float]]] = defaultdict(list)

    changed = True
    while changed:
        changed = False
        for n in list(sub_adj.keys()):
            if n in keep_nodes:
                continue
            nbrs = sub_adj.get(n, [])
            unique = list(dict.fromkeys(nbrs))
            if len(unique) != 2:
                continue
            a, b = unique
            if a == b:
                continue
            if b in sub_adj.get(a, []):
                continue
            ang1 = _angle(a, n); ang2 = _angle(n, b)
            diff = math.degrees(abs(((ang2 - ang1 + math.pi) % (2 * math.pi)) - math.pi))
            l_an = sub_edge_len.get((min(a, n), max(a, n)), math.hypot(n[0] - a[0], n[1] - a[1]))
            l_nb = sub_edge_len.get((min(n, b), max(n, b)), math.hypot(b[0] - n[0], b[1] - n[1]))
            should_merge = False
            if diff <= COLLINEAR_TOL_DEG:
                should_merge = True
            elif diff <= ELBOW_MERGE_TOL_DEG and (l_an + l_nb) <= 2 * SHORT_SEG_MM:
                should_merge = True
            if not should_merge:
                continue
            new_len = l_an + l_nb
            new_key = (min(a, b), max(a, b))
            prior_elbows: list[tuple[tuple[float, float], float]] = []
            for k_old in [(min(a, n), max(a, n)), (min(n, b), max(n, b))]:
                if k_old in edge_elbows:
                    prior_elbows.extend(edge_elbows.pop(k_old))
            if diff > COLLINEAR_TOL_DEG:
                prior_elbows.append((n, diff))
            if prior_elbows:
                edge_elbows[new_key] = prior_elbows
            sub_adj[a] = [x for x in sub_adj[a] if x != n] + [b]
            sub_adj[b] = [x for x in sub_adj[b] if x != n] + [a]
            del sub_adj[n]
            sub_edge_len.pop((min(a, n), max(a, n)), None)
            sub_edge_len.pop((min(n, b), max(n, b)), None)
            sub_edge_len[new_key] = new_len
            changed = True

    # merged edges 재구성
    merged_edges: list[tuple[tuple[float, float], tuple[float, float], float]] = []
    seen_keys: set = set()
    for n, nbrs in sub_adj.items():
        for m in nbrs:
            key = (min(n, m), max(n, m))
            if key in seen_keys:
                continue
            seen_keys.add(key)
            L = sub_edge_len.get(key, math.hypot(m[0] - n[0], m[1] - n[1]))
            merged_edges.append((n, m, L))
    merged_nodes = sorted(sub_adj.keys())

    _pcb(1.0, "선정 완료")
    return SelectionResult(
        source_pos=src,
        source_kind=src_kind,
        heads=selected_heads,
        distances=distances,
        edges=merged_edges,
        nodes_in_subgraph=merged_nodes,
        elbow_fittings={k: v for k, v in edge_elbows.items() if v},
        source_bridge_dist_mm=src_bridge_dist_mm,
        source_fallback=src_fallback,
    )


class _AnchorWindow:
    """anchored 작업창 W(§1) — convex_hull(head_region ∪ {alarm_xy}) 을
    ANCHOR_W_MARGIN_MM 만큼 팽창한 다각형. contains(pt) 프로토콜 제공.

    head_region 은 정점 리스트(pts)를 노출해야 한다(→W4 HeadRegion 이 정식화).
    팽창 판정은 "hull 내부 or hull 경계까지 거리 ≤ margin" — 순수 파이썬
    (shapely 미의존, BLOCKED.md #1).
    """

    def __init__(self, region_pts: list, alarm_xy: tuple[float, float],
                 margin_mm: float = ANCHOR_W_MARGIN_MM):
        pts = [(float(p[0]), float(p[1])) for p in region_pts]
        pts.append((float(alarm_xy[0]), float(alarm_xy[1])))
        self.hull = self._convex_hull(pts)
        self.margin = float(margin_mm)

    @staticmethod
    def _convex_hull(pts: list) -> list:
        """Andrew monotone chain — CCW hull 정점 리스트."""
        pts = sorted(set(pts))
        if len(pts) <= 2:
            return pts

        def cross(o, a, b):
            return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

        lower: list = []
        for p in pts:
            while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
                lower.pop()
            lower.append(p)
        upper: list = []
        for p in reversed(pts):
            while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
                upper.pop()
            upper.append(p)
        return lower[:-1] + upper[:-1]

    def contains(self, pt) -> bool:
        x, y = float(pt[0]), float(pt[1])
        h = self.hull
        if len(h) < 3:
            return any(math.hypot(x - p[0], y - p[1]) <= self.margin for p in h)
        inside = False
        for i in range(len(h)):
            x1, y1 = h[i]
            x2, y2 = h[(i + 1) % len(h)]
            if (y1 > y) != (y2 > y):
                xin = (x2 - x1) * (y - y1) / (y2 - y1) + x1
                if x < xin:
                    inside = not inside
        if inside:
            return True
        return any(
            _point_to_segment_dist(x, y, h[i][0], h[i][1],
                                   h[(i + 1) % len(h)][0], h[(i + 1) % len(h)][1]) <= self.margin
            for i in range(len(h))
        )


def bridge_targeted(
    graph: dict,
    edge_len: dict,
    source: tuple[float, float],
    accepted_heads: list,
    tols,
    bridge_edges_out: set | None = None,
    audit: dict | None = None,
    within=None,
) -> int:
    """anchored 표적 브릿지(W3) — 전역 ``_bridge_components`` 계단식 대체.

    봉합 후보 쌍을 comp(source) ↔ {region 내 승인 헤드 보유 컴포넌트} 로만 한정.
    tol 계단(도면비례)과 병합 후 재평가 루프는 전역 브릿지와 동일 원리이나,
    헤드 게이트를 통과하지 못한 컴포넌트(인접 세대망·범례·노이즈 조각)는 어떤
    tol 에서도 봉합되지 않는다 — 세대 간 가짜 봉합 차단.

    audit 지정 시 bridge 마다 p1(comp(source) 쪽)·p2(헤드 컴포넌트 쪽)·len_mm·tol·
    p1_in_source_comp 를 기록. p1 소속은 봉합 시점의 comp(source) 성장 이력 기준이라
    "모든 bridge 양단이 comp(source) 성장 이력에 속함"을 audit 로 검증 가능.
    within: 지정 시 봉합 양단 모두 이 영역(작업창 W) 안이어야 한다 — 세대를
        가로지르는 대형 컴포넌트가 W 밖 지점에서 봉합돼 동측 우회 경로가 생기는
        것을 차단(앵커가 봉합 방향을 유도한다는 §0 설계 원칙).
    반환: 추가한 bridge 수.
    """
    total = 0
    for tol in tols:
        tol2 = tol * tol
        while True:  # 병합 후 재평가 — comp(source) 가 자라며 사거리가 좁혀진다
            comps = _connected_components(graph)
            comp_of = {n: i for i, c in enumerate(comps) for n in c}
            sc = comp_of.get(source)
            if sc is None:
                return total
            src_comp = comps[sc]
            head_cids: set = set()
            for h in accepted_heads:
                pos = h.pos if hasattr(h, "pos") else (float(h[0]), float(h[1]))
                near = _nearest_graph_node(graph, pos)
                if near is None:
                    continue
                if math.hypot(pos[0] - near[0], pos[1] - near[1]) <= HEAD_BRIDGE_MAX_MM:
                    cid = comp_of[near]
                    if cid != sc:
                        head_cids.add(cid)
            if not head_cids:
                return total  # 승인 헤드 보유 컴포넌트 전부 병합 완료
            src_nodes = (list(src_comp) if within is None
                         else [v for v in src_comp if within.contains(v)])
            best = None
            bestd2 = tol2
            for cid in head_cids:
                cand = (comps[cid] if within is None
                        else [u for u in comps[cid] if within.contains(u)])
                for u in cand:
                    for v in src_nodes:
                        dx = u[0] - v[0]; dy = u[1] - v[1]
                        d2 = dx * dx + dy * dy
                        if d2 <= bestd2 and (best is None or d2 < bestd2):
                            bestd2 = d2
                            best = (u, v)
            if best is None:
                break  # 이 tol 로는 더 못 잇는다 — 다음 tol 로 완화
            u, v = best
            d = math.hypot(u[0] - v[0], u[1] - v[1])
            graph[u].add(v); graph[v].add(u)
            key = (min(u, v), max(u, v))
            edge_len[key] = d
            if bridge_edges_out is not None:
                bridge_edges_out.add(key)
            if audit is not None:
                # layers: 그래프 수준(좌표·길이만)에선 원 entity layer 가 소실돼
                # 복원 불가 — null 기록 (BLOCKED.md #7)
                audit.setdefault("bridges", []).append({
                    "p1": [v[0], v[1]], "p2": [u[0], u[1]],
                    "len_mm": d, "tol": tol, "layers": None,
                    "p1_in_source_comp": True,
                })
            total += 1
    return total


def collect_spatial_reselect_segments(dxf_path, layer_categories: dict[str, str],
                                      window) -> list[dict]:
    """W5 — S140 조건부 재선별의 공간 한정 실시예.

    1차 명목 수집(filter_pipenet_only)이 region 승인 헤드 도달망을 만들지 못한
    경우에만 호출된다(anchored 전용, 플래그 기본 off). 작업창 W(window) **내부**의
    OTHER 카테고리 레이어에서 LINE/ARC/LWPOLYLINE 만 저-prior 후보로 승인한다.

    SPLINE/ELLIPSE/3DFACE/DIMENSION/HATCH/닫힌 PL(CLOSED_PL_TOL_MM)은 음성
    유형 — 승인 금지. 파서(parse_dxf_bundle)는 SPLINE/ELLIPSE 를 PL 로 평탄화해
    entity dict 수준에서 원 유형을 구분할 수 없으므로(태그 추가는 골든
    entities_sig 를 깨 비트동일 위반), DXF 원본을 ezdxf 로 직접 스캔해 실제
    dxftype 으로 판정한다(BLOCKED.md #5).

    반환: pipe_entities 형식 dict 리스트 — 모든 정점이 W 안인 segment 만.
    """
    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()
    out: list[dict] = []
    for e in msp:
        layer = str(getattr(e.dxf, "layer", "") or "")
        if layer_categories.get(layer, "OTHER") != "OTHER":
            continue
        et = e.dxftype()
        if et == "LINE":
            p1 = (float(e.dxf.start.x), float(e.dxf.start.y))
            p2 = (float(e.dxf.end.x), float(e.dxf.end.y))
            if window.contains(p1) and window.contains(p2):
                out.append({"t": "L", "l": layer,
                            "p": [p1[0], p1[1], p2[0], p2[1]]})
        elif et == "ARC":
            cx, cy = float(e.dxf.center.x), float(e.dxf.center.y)
            r = float(e.dxf.radius)
            sa, ea = float(e.dxf.start_angle), float(e.dxf.end_angle)
            pa = (cx + r * math.cos(math.radians(sa)),
                  cy + r * math.sin(math.radians(sa)))
            pb = (cx + r * math.cos(math.radians(ea)),
                  cy + r * math.sin(math.radians(ea)))
            if window.contains(pa) and window.contains(pb):
                out.append({"t": "A", "l": layer, "c": [cx, cy], "r": r,
                            "a": [sa, ea]})
        elif et == "LWPOLYLINE":
            pts = [(float(p[0]), float(p[1])) for p in e.get_points()]
            if len(pts) < 2:
                continue
            closed = bool(e.closed) or (
                len(pts) >= 3
                and math.hypot(pts[0][0] - pts[-1][0],
                               pts[0][1] - pts[-1][1]) <= CLOSED_PL_TOL_MM)
            if closed:
                continue  # 닫힌 PL = 심볼/외곽선 — 음성 유형
            if all(window.contains(p) for p in pts):
                out.append({"t": "PL", "l": layer,
                            "p": [[p[0], p[1]] for p in pts]})
        # 그 외 dxftype (SPLINE/ELLIPSE/3DFACE/DIMENSION/HATCH 등) — 음성 유형
    return out


def select_worst30_heads_anchored(
    pipe_entities: list[dict],
    layer_categories: dict[str, str],
    alarm_xy: tuple[float, float],
    head_region,
    k: int = 30,
    manual_heads: list[tuple[float, float]] | None = None,
    zones: list[tuple[float, float, float, float]] | None = None,
    branch_zones: "list[tuple[float, float, float, float]] | HeadRegion | None" = None,
    audit_out: dict | None = None,
    progress_cb=None,
    spatial_reselect: bool = False,
    tee_split: bool = False,
    dxf_path=None,
) -> SelectionResult:
    """2앵커 anchored 선정(§1 계약) — ``alarm_xy``·``head_region`` 필수.

    비-anchored 경로(select_worst30_heads)는 손대지 않고 보존(골든 불변) — 신규
    로직은 전부 이 anchored 전용 함수에서만 발동한다. 순서가 다르다:
      ① 헤드를 region 게이트(W1) 후 먼저 확정
      ② 소스를 attach_source(W2)로 헤드 보유 컴포넌트에 결합 (blind nearest 금지)
      ③ 전역 _bridge_components 계단식 대신 표적 브릿지(W3)
      ④ corridor 제한·SPT·후반부는 기존 경로와 동일 primitive 공유
    force_connect(무제한 봉합)는 이 경로에서 호출 금지 — 기계실(탱크) 추출 전용.
    audit_out: 지정 시 source_attach/bridges/heads(unreachable 포함) 근거 기록(→W7).
    spatial_reselect: W5 공간한정 조건부 재선별 플래그(기본 off). 1차 명목 수집
        결과로 region 승인 헤드 도달망 구성에 실패한 경우에만 발동하며,
        dxf_path(원본 DXF) 가 필요하다.
    tee_split: T분기 edge-split 플래그(기본 off). weld 이전에 실행돼 weld 가
        발명해야 할 연결을 줄인다(→audit tee_splits).
    """
    if alarm_xy is None:
        raise ValueError("anchored: alarm_xy(수동 알람밸브 좌표) 필수")
    if head_region is None:
        raise ValueError("anchored: head_region(헤드 영역 다각형) 필수")
    audit = audit_out if audit_out is not None else {}
    _pcb = progress_cb if callable(progress_cb) else (lambda f, m: None)
    # 작업창 W (§1) — W3 브릿지 양단 한정·W5 재선별 공간 한정에 공용
    _region_pts = getattr(head_region, "pts", None)
    _W = (_AnchorWindow(_region_pts, (float(alarm_xy[0]), float(alarm_xy[1])))
          if _region_pts else None)
    # W6 — 호출측이 build_input_tables(anchor_window=...) 에 그대로 전달할 수 있게
    # 작업창을 노출 (객체 — JSON 직렬화 대상 아님, W7 스키마 밖).
    audit["anchor_window"] = _W
    # ① 헤드 — region 게이트(W1) 통과한 최종 승인 후보만 (그래프와 무관 — 선확정)
    if manual_heads is not None:
        heads = [HeadCandidate(pos=_round_pt(x, y), raw=(x, y), block_name="(user)", layer="_user")
                 for x, y in manual_heads if head_region.contains((x, y))]
    else:
        heads = _find_head_candidates(pipe_entities, layer_categories, region=head_region)
    if zones:
        heads = [h for h in heads if _point_in_zones(h.pos[0], h.pos[1], zones)]
    if len(heads) < k:
        k = len(heads)
    audit.setdefault("heads", {})["detected_in_region"] = len(heads)
    _pcb(0.0, "배관망 그래프 구성 중")
    _idx = _NodeIndex()
    graph, edge_len = _build_graph(pipe_entities, node_index=_idx,
                                   layer_categories=layer_categories)
    # ── W5: 공간한정 조건부 재선별 (S140 조건부 재선별의 공간 한정 실시예) ──
    # 발동 조건: 플래그 on **그리고** 1차 명목 수집(filter_pipenet_only) 그래프가
    # region 승인 헤드에 도달하지 못할 때만 (헤드 최근접 노드 없음/HEAD_BRIDGE_MAX 밖).
    if spatial_reselect and heads and _W is not None and dxf_path is not None:
        _attachable = any(
            (_nn := _nearest_graph_node(graph, h.pos)) is not None
            and math.hypot(_nn[0] - h.pos[0], _nn[1] - h.pos[1]) <= HEAD_BRIDGE_MAX_MM
            for h in heads)
        if not _attachable:
            _pcb(0.03, "공간한정 조건부 재선별 중")
            resel = collect_spatial_reselect_segments(dxf_path, layer_categories, _W)
            g2, el2 = _build_graph(resel, node_index=_idx, layer_categories=None)
            new_keys = [kk for kk in el2 if kk not in edge_len]
            for u, nbs in g2.items():
                graph[u] |= nbs
            for kk, _L in el2.items():
                prev = edge_len.get(kk)
                if prev is None or _L < prev:
                    edge_len[kk] = _L
            # 승인된 비명목 edge 태깅 → audit 점유율 집계(→W7 nonnominal)
            _nn_len = float(sum(el2[kk] for kk in new_keys))
            _total_len = float(sum(edge_len.values()))
            audit["nonnominal"] = {
                "edge_count": len(new_keys),
                "len_mm": _nn_len,
                "ratio": (_nn_len / _total_len) if _total_len else 0.0,
            }
    collapse_parallel_ladders(graph, edge_len)
    # T분기 복원은 weld 이전 — 주배관 중간에 닿는 가지관을 제자리에서 접속시켜
    # weld 가 엉뚱한 조각/끝점에 붙이는 오접합을 애초에 없앤다.
    if tee_split:
        _tee: list = []
        _split_tee_branches(graph, edge_len, splits_out=_tee)
        _syms = [(en["p"][0], en["p"][1]) for en in pipe_entities
                 if en["t"] == "I"
                 and layer_categories.get(en.get("l", ""), "OTHER") != "HEAD"]
        for rec in _tee:
            rec["sym_mm"] = min((math.hypot(rec["p"][0] - sx, rec["p"][1] - sy)
                                 for sx, sy in _syms), default=None)
        audit["tee_splits"] = _tee
    _diag = _graph_diag(graph)
    _penalty_keys: set = set()
    _weld_keys: set = set()  # W7 — weld 만 별도 수집 (audit welds 항목)
    _weld_dangling_endpoints(graph, edge_len,
                             weld_tol=_adaptive_weld_tol(_diag),
                             weld_cone_deg=_WELD_CONE_DEG,
                             weld_edges_out=_weld_keys)
    _penalty_keys |= _weld_keys
    audit["welds"] = [{"p1": [kk[0][0], kk[0][1]], "p2": [kk[1][0], kk[1][1]],
                       "len_mm": edge_len.get(kk, 0.0)} for kk in sorted(_weld_keys)]
    _pcb(0.05, "평행 배관 정리 완료")
    _pcb(0.1, "알람밸브 결합 중")
    # ② 소스 — attach_source(W2). 헤드 보유 컴포넌트 우선, blind nearest 금지.
    src_raw = _round_pt(float(alarm_xy[0]), float(alarm_xy[1]))
    comps = _connected_components(graph)
    comp_of = {n: i for i, c in enumerate(comps) for n in c}
    src, attach_key = attach_source(src_raw, graph, comp_of, heads, edge_len, audit)
    if attach_key is not None:
        _penalty_keys.add(attach_key)  # source 접속선도 추정연결 — 라우팅 penalty
    src_bridge_dist_mm = float(audit["source_attach"]["dist_mm"] or 0.0)
    # ③ 표적 브릿지(W3) — tol 계단은 비-anchored 와 동일(도면 스케일 비례).
    #    브릿지 양단은 작업창 W 안으로 한정 — 앵커가 봉합 방향을 유도(§0).
    #    도면 현실상 배관 컴포넌트가 세대 경계를 넘어 이어지므로(예: 대명동
    #    comp12, x 252k→262k) W 제한 없이는 동측 지점에서 봉합될 수 있다.
    bridge_targeted(graph, edge_len, src, heads, _adaptive_bridge_tols(_diag),
                    bridge_edges_out=_penalty_keys, audit=audit, within=_W)
    _pcb(0.7, "표적 브릿지 완료")
    # 미도달 헤드 보고(W1.2) — 조용한 drop 금지: 추출 결함 신호로 audit 에 기록
    audit["heads"]["unreachable"] = [
        [p[0], p[1]] for p in find_unreachable_region_heads(graph, src, heads)
    ]
    audit["heads"]["attached"] = len(heads) - len(audit["heads"]["unreachable"])
    # 급수 감사 — corridor 제한·SPT 이전(사이클 보유 상태)에 재야 루프/드레인이 보인다.
    audit["water"] = water_load_audit(graph, edge_len, src, heads,
                                      penalty_keys=_penalty_keys)
    # ④ corridor 제한 → SPT → 공통 후반부 (기존 primitive 그대로)
    region_nodes = _restrict_to_branch_region(graph, edge_len, src, branch_zones,
                                              penalty_keys=_penalty_keys)
    # W7 — corridor 집계: 제한 후 영역 밖에 남은 edge = source→영역 corridor
    # (분기영역 미지정/no-op 이면 0). 별도 경로 탐색 없음 — 제한 결과 재사용.
    if region_nodes:
        _cor_keys: set = set()
        for _u, _nbs in graph.items():
            for _v in _nbs:
                if _u not in region_nodes or _v not in region_nodes:
                    _cor_keys.add((min(_u, _v), max(_u, _v)))
        _cor_nodes = {n for kk in _cor_keys for n in kk}
        audit["corridor"] = {"node_count": len(_cor_nodes),
                             "len_mm": float(sum(edge_len.get(kk, 0.0)
                                                 for kk in _cor_keys))}
    else:
        audit["corridor"] = {"node_count": 0, "len_mm": 0.0}
    force_spanning_tree(graph, edge_len, source=src, penalty_keys=_penalty_keys)
    _hd_keys: set = set()
    res = _finalize_selection(graph, edge_len, src, "manual_anchored", heads, k, _pcb,
                              src_bridge_dist_mm, False, head_drop_out=_hd_keys)
    audit["head_drops"] = [{"p1": [kk[0][0], kk[0][1]], "p2": [kk[1][0], kk[1][1]],
                            "len_mm": edge_len.get(kk, 0.0)} for kk in sorted(_hd_keys)]
    # W7 — 축적 dict → 정식 스키마. 반환값에 포함 (r30_combined 가 JSON 직렬화)
    res.audit = ExtractionAudit.from_audit_dict(audit)
    return res


# ────────────────────────────────────────────────────────────────────────────
# 3) Stage 3 — Input 5 tables + Meta
# ────────────────────────────────────────────────────────────────────────────


@dataclass(slots=True)
class PipeTables:
    nodes: list[dict] = field(default_factory=list)      # [{label,elevation,io_node,x,y}]
    pipes: list[dict] = field(default_factory=list)      # [{label,in,out,type,dia,length,elev,c,status,group}]
    nozzles: list[dict] = field(default_factory=list)    # [{label,in,out,status,lib,flow_m3s,flow_lmin}]
    fittings: list[dict] = field(default_factory=list)   # [{pipe,in,out,type,count}]
    equipment: list[dict] = field(default_factory=list)  # [{pipe,in,out,label,desc,eq_len,rel_pos}]
    meta: list[tuple[str, str]] = field(default_factory=list)

    def as_dict(self) -> dict:
        """JSON 직렬화용 dict. job state 저장 / SSE 이벤트 payload 에 사용.

        meta 는 tuple 리스트라 JSON round-trip 시 list 로 바뀌므로 from_dict 에서 복원.
        """
        return {
            "nodes": self.nodes, "pipes": self.pipes, "nozzles": self.nozzles,
            "fittings": self.fittings, "equipment": self.equipment,
            "meta": [list(m) for m in self.meta],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "PipeTables":
        """as_dict() 역변환. meta 는 (key, value) tuple 로 되돌린다."""
        return cls(
            nodes=list(d.get("nodes") or []),
            pipes=list(d.get("pipes") or []),
            nozzles=list(d.get("nozzles") or []),
            fittings=list(d.get("fittings") or []),
            equipment=list(d.get("equipment") or []),
            meta=[tuple(m) for m in (d.get("meta") or [])],
        )


# 배관 재질 — DXF 에 없는 설계 정보이므로 자동 분류하지 않는다. 기본은 강관(KSD 3507,
# C=120). 사용자가 평면도에서 지정한 영역(zone=단위세대 내부) 안의 배관만 CPVC(C=150)로
# 유지한다. 세대 배관은 통상 CPVC, 간선/입상관은 강관이라는 현장 관행을 사용자 영역
# 지정으로 표현. C-factor(=roughness-or-c)가 PIPENET 마찰손실에 직접 반영되는 재질값.


def _point_in_zones(px: float, py: float,
                    zones: list[tuple[float, float, float, float]] | None) -> bool:
    """점이 zone(축정렬 사각형) union 안에 있으면 True. 좌표 순서 무관."""
    if not zones:
        return False
    for (zx1, zy1, zx2, zy2) in zones:
        lo_x, hi_x = (zx1, zx2) if zx1 <= zx2 else (zx2, zx1)
        lo_y, hi_y = (zy1, zy2) if zy1 <= zy2 else (zy2, zy1)
        if lo_x <= px <= hi_x and lo_y <= py <= hi_y:
            return True
    return False


# 가지배관 직각화 각도 임계값(deg) — 가지 edge 가 축(0/90°)에서 이 이내면 축정렬로
# 스냅, 45° 근방(진짜 대각선)은 실좌표 유지. 표시 전용(length_mm·연결 불변).


def _classify_branch_edges(edges, head_points, source_point):
    """소스 기준 트리에서 '가지배관' edge 를 판별한다.

    교차·주배관(cross main)은 가지선을 여러 갈래로 분기시키는 spine 이고, 가지배관은
    그 끝에 헤드가 달린 (분기 없는) 열이다. 규칙: 배관(부모 u→자식 v)의 하류 서브트리에
    splitter(헤드를 가진 자식이 2개 이상인 노드)가 하나도 없으면 v 쪽은 가지선 → 가지배관.
    splitter 가 있으면 교차배관.

    반환: (branch_edge_keys: set[frozenset{ka,kb}], key_fn) — 가지 edge 키 집합.
    head_points/source_point 가 없으면 빈 집합(=전부 비가지).
    """
    def _key(p):
        return (round(float(p[0]), 3), round(float(p[1]), 3))
    if source_point is None:
        return set(), _key
    from collections import deque as _deque
    adj: dict = defaultdict(set)
    for e in edges:
        ka, kb = _key(e[0]), _key(e[1])
        if ka != kb:
            adj[ka].add(kb); adj[kb].add(ka)
    root = _key(source_point)
    if root not in adj:
        return set(), _key
    parent: dict = {root: None}
    order: list = [root]
    q = _deque([root])
    while q:
        u = q.popleft()
        for v in adj[u]:
            if v not in parent:
                parent[v] = u; order.append(v); q.append(v)
    children: dict = defaultdict(list)
    for v, p in parent.items():
        if p is not None:
            children[p].append(v)
    head_set = {_key(h) for h in (head_points or [])}
    dh: dict = {}
    for v in reversed(order):
        c = 1 if v in head_set else 0
        for ch in children[v]:
            c += dh[ch]
        dh[v] = c
    # splitter 서브트리 판정 — 하류에서 헤드열이 2갈래+로 갈라지면 교차배관.
    sub_split: dict = {}
    for v in reversed(order):
        head_bearing_kids = sum(1 for ch in children[v] if dh[ch] >= 1)
        s = head_bearing_kids >= 2
        for ch in children[v]:
            s = s or sub_split[ch]
        sub_split[v] = s
    # ── 교차배관 spine(trunk) 추적: 소스에서 방향 연속성으로 직진하는 경로.
    #    splitter 규칙만으론 마지막 분기 이후의 교차배관 tail 을 가지선과 구분 못 하므로
    #    (tail 서브트리도 splitter 가 없음), 진행방향이 이어지는 간선을 trunk 로 표시한다.
    TRUNK_TURN_TOL = math.radians(45.0)  # 45° 이상 꺾이면 trunk 종료(가지 진입).

    def _dir(u, v):
        return math.atan2(v[1] - u[1], v[0] - u[0])

    def _angdiff(a, b):
        return abs((a - b + math.pi) % (2.0 * math.pi) - math.pi)

    trunk: set = set()
    cur, inc = root, None
    while True:
        kids = children.get(cur, [])
        if not kids:
            break
        if inc is None:
            nxt = max(kids, key=lambda c: dh[c])       # 소스: 헤드 최다 방향으로 출발
        else:
            nxt = min(kids, key=lambda c: _angdiff(_dir(cur, c), inc))
            if _angdiff(_dir(cur, nxt), inc) > TRUNK_TURN_TOL:
                break                                   # 크게 꺾임 → trunk 끝(가지 시작)
        trunk.add(frozenset((cur, nxt)))
        inc = _dir(cur, nxt)
        cur = nxt
    branch: set = set()
    for v in order:
        u = parent[v]
        if u is None:
            continue
        fe = frozenset((u, v))
        if fe in trunk:               # 교차배관 spine — 고정
            continue
        if sub_split.get(v, True):    # 하류 분기 존재 — 교차배관
            continue
        branch.add(fe)                # 분기 없는 헤드열 = 가지배관
    return branch, _key


def orthogonalize_edge_positions(edges, *, head_points=None, source_point=None,
                                 tol_deg: float = ORTHO_SNAP_TOL_DEG) -> dict:
    """가지배관만 직각에 스냅한 표시좌표맵을 반환(표시 전용, 교차·주배관/소스 고정).

    head_points·source_point 가 주어지면 트리에서 가지배관을 판별해(_classify_branch_edges)
    가지 edge 만 축정렬 제약을 건다. 교차배관·소스 노드는 실 DXF 좌표에 고정(anchor).
      · 근축 수평 가지 → 두 끝점 같은 Y, 근축 수직 가지 → 같은 X (X·Y 독립)
      · 각 축-성분: 고정 노드 1개면 그 좌표로 정렬(tee 앵커), 0개면 평균, ≥2개면
        충돌이므로 이동 안 함(실 DXF 유지).
    45° 근방 가지·비가지 배관은 실좌표 유지. length_mm(유압 권위값) 불변.

    head/source 미지정이면 모든 근축 edge 스냅(구 동작, 폴백).
    반환: {(rx, ry): (x', y')}. 매칭 실패 시 호출측이 원좌표 사용.
    """
    branch_keys, _key = _classify_branch_edges(edges, head_points, source_point)
    branch_only = bool(branch_keys) or (source_point is not None and head_points is not None)

    orig: dict = {}
    for e in edges:
        a, b = e[0], e[1]
        orig.setdefault(_key(a), (float(a[0]), float(a[1])))
        orig.setdefault(_key(b), (float(b[0]), float(b[1])))
    if not edges:
        return dict(orig)

    tol = math.radians(tol_deg)
    parent_x = {k: k for k in orig}
    parent_y = {k: k for k in orig}

    def _find(par, k):
        root = k
        while par[root] != root:
            root = par[root]
        while par[k] != root:
            par[k], k = root, par[k]
        return root

    def _union(par, i, j):
        ri, rj = _find(par, i), _find(par, j)
        if ri != rj:
            par[ri] = rj

    # 고정 노드 = 비가지(교차·주배관) edge 에 닿는 노드 + 소스. branch_only 아니면 고정 없음.
    fixed: set = set()
    if branch_only:
        for e in edges:
            ka, kb = _key(e[0]), _key(e[1])
            if ka == kb:
                continue
            if frozenset((ka, kb)) not in branch_keys:
                fixed.add(ka); fixed.add(kb)
        if source_point is not None:
            fixed.add(_key(source_point))

    for e in edges:
        a, b = e[0], e[1]
        ka, kb = _key(a), _key(b)
        if ka == kb:
            continue
        ang = math.atan2(float(b[1]) - float(a[1]), float(b[0]) - float(a[0])) % math.pi
        d_h = min(ang, math.pi - ang)      # 수평축(0/π)까지 각거리
        d_v = abs(ang - math.pi / 2.0)     # 수직축(π/2)까지 각거리
        if branch_only:
            if frozenset((ka, kb)) not in branch_keys:
                continue  # 교차·주배관 — 스냅 안 함(실 DXF 고정)
            # 가지배관은 가까운 축으로 무조건 정렬(0/90/180/270). tol 게이트 없음.
            if d_h <= d_v:
                _union(parent_y, ka, kb)   # 수평 → 같은 Y
            else:
                _union(parent_x, ka, kb)   # 수직 → 같은 X
        elif d_h <= tol and d_h <= d_v:
            _union(parent_y, ka, kb)       # (폴백) 근축 수평 → 같은 Y
        elif d_v <= tol:
            _union(parent_x, ka, kb)       # (폴백) 근축 수직 → 같은 X
        # else: 대각선 — 스냅 안 함

    def _solve(par, coord_idx):
        groups: dict = defaultdict(list)
        for k in orig:
            groups[_find(par, k)].append(k)
        out: dict = {}
        for _r, members in groups.items():
            fixed_vals = [orig[k][coord_idx] for k in members if k in fixed]
            if len(fixed_vals) >= 2 and (max(fixed_vals) - min(fixed_vals)) > 1e-6:
                target = None  # 고정 노드 충돌 → 이동 안 함(각자 실좌표)
            elif fixed_vals:
                target = fixed_vals[0]      # tee 앵커
            else:
                target = sum(orig[k][coord_idx] for k in members) / len(members)
            for k in members:
                out[k] = orig[k][coord_idx] if target is None else target
        return out

    sx = _solve(parent_x, 0)
    sy = _solve(parent_y, 1)
    result = {k: (sx[k], sy[k]) for k in orig}
    if branch_only and branch_keys:
        result = _separate_overlapping_branches(
            result, edges, branch_keys, fixed, _key)
    return result


def _separate_overlapping_branches(result, edges, branch_keys, fixed, key_fn,
                                   *, min_overlap: float = 150.0,
                                   max_passes: int = 6) -> dict:
    """표시 전용: 서로 다른 가지 컴포넌트가 같은 축선에 겹쳐(루프처럼 보임) 그려지면
    이동 가능한(비고정) 노드를 겹침 축의 수직 방향으로 한 레인만큼 밀어 해소한다.

    고정(교차·주배관/소스) 노드는 절대 이동하지 않는다 → 고정 tee 로 연결되는 edge 는
    짧은 사선(가지 분기 표시)이 될 수 있으나 병렬 드롭 두 줄이 각자 레인을 가져 겹침이
    사라진다. length_mm(유압)은 raw 좌표로 계산되므로 불변.
    """
    from collections import deque as _deque
    adj: dict = defaultdict(set)
    for e in edges:
        ka, kb = key_fn(e[0]), key_fn(e[1])
        if ka == kb:
            continue
        if frozenset((ka, kb)) in branch_keys:
            adj[ka].add(kb); adj[kb].add(ka)
    if not adj:
        return result

    branch_edge_list = []
    for e in edges:
        ka, kb = key_fn(e[0]), key_fn(e[1])
        if ka == kb or frozenset((ka, kb)) not in branch_keys:
            continue
        branch_edge_list.append((ka, kb))
    if not branch_edge_list:
        return result

    # 각 컴포넌트를 고정 노드(교차·주배관 tee) 기준으로 뿌리내려 부모/자식·깊이를 구한다.
    parent: dict = {}
    depth: dict = {}
    for n in adj:
        if n in parent:
            continue
        comp = []
        stack = [n]; seen_c: set = set()
        while stack:
            u = stack.pop()
            if u in seen_c:
                continue
            seen_c.add(u); comp.append(u)
            stack.extend(adj[u] - seen_c)
        root = next((k for k in comp if k in fixed), comp[0])
        parent[root] = None; depth[root] = 0
        q = _deque([root])
        while q:
            u = q.popleft()
            for v in adj[u]:
                if v not in parent:
                    parent[v] = u; depth[v] = depth[u] + 1; q.append(v)

    # deep 노드(뿌리에서 먼 쪽)를 루트로 하는 서브트리(자유 노드만) — 이동 대상 후보.
    def _subtree_free(deep, blocked):
        out = []
        stack = [deep]; local: set = {blocked}
        while stack:
            u = stack.pop()
            if u in local:
                continue
            local.add(u)
            if u not in fixed:
                out.append(u)
            for v in adj[u]:
                if v not in local:
                    stack.append(v)
        return out

    result = dict(result)
    seg_lens = sorted(
        math.hypot(result[kb][0] - result[ka][0], result[kb][1] - result[ka][1])
        for ka, kb in branch_edge_list)
    lane = min(max(seg_lens[len(seg_lens) // 2] * 0.5, 250.0), 1200.0)

    def _collisions() -> list:
        segs = []  # (edge_idx, orient, line, lo, hi)
        for ei, (ka, kb) in enumerate(branch_edge_list):
            (ax, ay), (bx, by) = result[ka], result[kb]
            if abs(ax - bx) < 1.0 and abs(ay - by) >= 1.0:
                segs.append((ei, 'V', (ax + bx) / 2, min(ay, by), max(ay, by)))
            elif abs(ay - by) < 1.0 and abs(ax - bx) >= 1.0:
                segs.append((ei, 'H', (ay + by) / 2, min(ax, bx), max(ax, bx)))
        cols = []
        for i in range(len(segs)):
            ei, oi, li, loi, hii = segs[i]
            for j in range(i + 1, len(segs)):
                ej, oj, lj, loj, hij = segs[j]
                if oi != oj or abs(li - lj) >= 1.0:
                    continue
                ov = min(hii, hij) - max(loi, loj)
                if ov > min_overlap:
                    cols.append((ei, ej, oi, ov))
        return cols

    def _mover_set(ei):
        """edge ei 의 deep 서브트리(자유 노드)를 이동 후보로 — 고정 노드가 섞이면 None."""
        ka, kb = branch_edge_list[ei]
        deep, anchor = (ka, kb) if depth.get(ka, 0) >= depth.get(kb, 0) else (kb, ka)
        if deep in fixed:
            return None
        return _subtree_free(deep, anchor)

    for _ in range(max_passes):
        cols = _collisions()
        if not cols:
            break
        cols.sort(key=lambda t: -t[3])
        ei, ej, axis, _ov = cols[0]
        cand_i, cand_j = _mover_set(ei), _mover_set(ej)
        # 더 작은(덜 앵커된) 서브트리를 민다. 둘 다 불가면 다음 후보로.
        options = [c for c in (cand_i, cand_j) if c]
        if not options:
            break
        move = min(options, key=len)
        idx = 0 if axis == 'V' else 1
        other = cand_j if move is cand_i else cand_i
        line = (sum(result[k][idx] for k in other) / len(other)) if other else \
            (result[branch_edge_list[ej if move is cand_i else ei][0]][idx])
        mcen = sum(result[k][idx] for k in move) / len(move)
        step = (1.0 if mcen >= line else -1.0) * lane
        for k in move:
            x, y = result[k]
            result[k] = (x + step, y) if idx == 0 else (x, y + step)
    return result


def build_input_tables(
    selection: SelectionResult,
    pipe_entities: list[dict] | None = None,
    *,
    project_title: str = "Remote 30 Prototype",
    cpvc_zones: list[tuple[float, float, float, float]] | None = None,
    anchor_window=None,
) -> PipeTables:
    """선정 결과 → 5 테이블. pipe_entities 가 있으면 FX(flexible) Equipment 도 추출.

    cpvc_zones: 이 영역(단위세대 내부) 안에 배관 중점이 들어오면 CPVC(C=150)로 표기.
    비어있으면 전 배관 강관(C=120).
    anchor_window: anchored 모드의 작업창 W(contains(pt) 노출 객체). 지정 시 관경
        텍스트 후보를 W 내부로 제한 — 범례 표(x≈288k)의 관경 문자 오염 차단(W6).
        None(비-anchored)이면 기존과 동일.
    """
    tables = PipeTables()
    if not selection.heads or selection.source_pos is None:
        return tables

    # 노드 라벨링 — 알람밸브 = 10, 나머지 1 부터 순차
    pos_to_label: dict[tuple[float, float], str] = {}
    label_to_pos: dict[str, tuple[float, float]] = {}
    counter = [10]

    def _label_node(pos: tuple[float, float]) -> str:
        if pos in pos_to_label:
            return pos_to_label[pos]
        lab = str(counter[0]); counter[0] += 1
        pos_to_label[pos] = lab
        label_to_pos[lab] = pos
        return lab

    src_label = _label_node(selection.source_pos)
    for n in selection.nodes_in_subgraph:
        _label_node(n)
    head_node_label: dict[tuple[float, float], str] = {}
    for h, dist in zip(selection.heads, selection.distances):
        snap = h.pos
        lab = _label_node(snap)
        head_node_label[snap] = lab

    # Nodes — 가지배관 표시좌표(x,y)만 직각화 스냅. 교차·주배관/소스는 실 DXF 고정.
    # length·diameter 는 아래에서 raw DXF 좌표로 계산하므로 유압 결과 불변(표시 전용).
    _ortho = orthogonalize_edge_positions(
        selection.edges,
        head_points=[h.pos for h in selection.heads],
        source_point=selection.source_pos)
    # de-overlap 표시 후 남는 가지배관 대각선 → L-벤드(직각 꺾임)로 그리기 위한 분류.
    _branch_keys, _bkey = _classify_branch_edges(
        selection.edges, [h.pos for h in selection.heads], selection.source_pos)

    def _oxy(p):
        return _ortho.get((round(float(p[0]), 3), round(float(p[1]), 3)),
                          (float(p[0]), float(p[1])))

    for label, pos in label_to_pos.items():
        io_node = "Input" if label == src_label else "No"
        ox, oy = _oxy(pos)
        tables.nodes.append({
            "label": label, "elevation": 2.8, "io_node": io_node,
            "x": int(round(ox)), "y": int(round(oy)),
        })

    # ====== Diameter 추론 — 3단계 알고리즘
    # ① DXF TEXT 패턴 5종 추출 (노이즈 워드 필터)
    # ② NFPC 103 별표 1 "가"칸 (폐쇄형 SP) — 담당 헤드 수 → 최소 호칭경 매핑
    # ③ 결정: 텍스트 매칭이 있으면 max(텍스트값, NFPC최소값), 없으면 NFPC fallback
    DIA_PATTERNS = [
        re.compile(r"\b(\d{2,3})\s*A\b"),                                # 25A
        re.compile(r"^\s*(\d{2,3})\s*$"),                                # 순수 숫자 (이 도면 dominant)
        re.compile(r"[Øø]\s*(\d{2,3})"),                                 # Ø25
        re.compile(r"DN\s*(\d{2,3})"),                                   # DN25
        re.compile(r"(?<![0-9])(\d{2,3})\s*mm(?![0-9])"),                # 25mm
    ]
    NOISE_KEYWORDS = ("호스", "방수구", "소화전", "옥내", "HOSE", "EA", "KG", "℃",
                       "SET", "SCALE", "PUMP", "펌프", "TANK", "탱크")
    VALID_DIA = {15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300}
    DIA_RANGE_LIMIT_MM = 1500.0  # 호칭경 텍스트는 보통 배관에 1.5m 이내 가까이 위치

    dia_text_pts: list[tuple[float, float, int]] = []  # (x, y, dia_mm)
    if pipe_entities:
        for en in pipe_entities:
            if en["t"] != "T":
                continue
            v = (en.get("v") or "").strip()
            if not v:
                continue
            if any(nw in v for nw in NOISE_KEYWORDS):
                continue  # 옥내소화전 / 헤드 라벨 / 스펙 표 등 노이즈
            # W6 — anchored 작업창 밖 관경 문자(범례 표 등) 후보 배제
            if anchor_window is not None and not anchor_window.contains(
                    (float(en["p"][0]), float(en["p"][1]))):
                continue
            for pat in DIA_PATTERNS:
                m = pat.search(v)
                if not m:
                    continue
                try:
                    d = int(m.group(1))
                except ValueError:
                    continue
                if d in VALID_DIA:
                    dia_text_pts.append((en["p"][0], en["p"][1], d))
                    break

    # ── NFPC 103 별표 1 "가" 칸 (폐쇄형 SP, 가장 일반)
    def _nfpc_min_bore_mm(head_count: int) -> int:
        if head_count <= 2:   return 25
        if head_count <= 3:   return 32
        if head_count <= 5:   return 40
        if head_count <= 10:  return 50
        if head_count <= 30:  return 65
        if head_count <= 60:  return 80
        if head_count <= 80:  return 90
        if head_count <= 100: return 100
        if head_count <= 160: return 125
        return 150

    # ── subgraph 안 src 부터의 BFS tree → pipe 별 downstream 헤드 수
    src_pos = selection.source_pos
    adj_sub: dict = defaultdict(list)
    for ea, eb, _ in selection.edges:
        adj_sub[ea].append(eb); adj_sub[eb].append(ea)
    parent_map: dict = {src_pos: None}
    bfs_q: list = [src_pos]
    while bfs_q:
        cur = bfs_q.pop(0)
        for nb in adj_sub[cur]:
            if nb not in parent_map:
                parent_map[nb] = cur
                bfs_q.append(nb)
    children_of: dict = defaultdict(list)
    for nd, pr in parent_map.items():
        if pr is not None:
            children_of[pr].append(nd)
    selected_head_set = {h.pos for h in selection.heads}
    subtree_count: dict = {}
    def _subtree_calc(n):
        cnt = 1 if n in selected_head_set else 0
        for c in children_of[n]:
            cnt += _subtree_calc(c)
        subtree_count[n] = cnt
        return cnt
    if src_pos is not None:
        _subtree_calc(src_pos)

    def _downstream_heads(a, b) -> int:
        if parent_map.get(b) == a: return subtree_count.get(b, 0)
        if parent_map.get(a) == b: return subtree_count.get(a, 0)
        return 0

    def _point_seg_dist(px, py, ax, ay, bx, by) -> float:
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        if L2 < 1e-9:
            return math.hypot(px - ax, py - ay)
        t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
        qx, qy = ax + t * dx, ay + t * dy
        return math.hypot(px - qx, py - qy)

    diameter_source_counter: dict[str, int] = {"text": 0, "nfpc_min": 0, "nfpc_fallback": 0}

    def _pipe_diameter(a: tuple[float, float], b: tuple[float, float]) -> int:
        nfpc_min = _nfpc_min_bore_mm(_downstream_heads(a, b))
        # 텍스트 매칭 — 점-선분 수직거리, 1500mm 이내
        best_text = None; best_d = DIA_RANGE_LIMIT_MM
        for tx, ty, dia in dia_text_pts:
            d = _point_seg_dist(tx, ty, a[0], a[1], b[0], b[1])
            if d < best_d:
                best_d = d; best_text = dia
        if best_text is None:
            diameter_source_counter["nfpc_fallback"] += 1
            return nfpc_min
        # 안전측: 텍스트 값이 별표 1 최소보다 작으면 별표 1 채택
        if best_text < nfpc_min:
            diameter_source_counter["nfpc_min"] += 1
            return nfpc_min
        diameter_source_counter["text"] += 1
        return best_text

    # Pipes + edge key → pipe label mapping
    edge_key_to_pipe: dict[tuple, str] = {}
    pipe_label_counter = 10
    cpvc_pipe_count = 0
    for a, b, length_mm in selection.edges:
        la = pos_to_label[a]; lb = pos_to_label[b]
        try:
            la_i, lb_i = int(la), int(lb)
            if la_i > lb_i:
                la, lb = lb, la
        except ValueError:
            pass
        plabel = str(pipe_label_counter)
        edge_key_to_pipe[(min(a, b), max(a, b))] = plabel
        dia = _pipe_diameter(a, b)
        # 배관 중점이 사용자 지정 CPVC 영역(단위세대) 안이면 CPVC, 아니면 강관.
        mx, my = (a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0
        if _point_in_zones(mx, my, cpvc_zones):
            ptype, cfac = CPVC_PIPE_TYPE, CPVC_C_FACTOR
            cpvc_pipe_count += 1
        else:
            ptype, cfac = STEEL_PIPE_TYPE, STEEL_C_FACTOR
        pipe_dict = {
            "label": plabel,
            "in": la, "out": lb,
            "type": ptype,
            "dia": dia,
            # PIPENET/K-solver 는 length=0 배관을 거부(특이행렬). 좌표가 거의 겹치는
            # 노드쌍(클러스터 잔여)은 round 후 0.0 이 되므로 10mm 하한 강제.
            "length": max(round(length_mm / 1000.0, 3), 0.01),
            "elev": 0.0,
            "c": cfac,
            "status": "Normal",
            "group": "Unset",
        }
        # 표시 전용 L-벤드 — _separate_overlapping_branches 가 가지 서브트리를 옆으로
        # 밀어 겹침을 없애면 고정 tee 로 이어지는 가지배관이 살짝 기운 대각선이 된다.
        # 그 대각선을 직각 꺾임(상류 쪽 짧은 다리 + 병렬 드롭)으로 표시해 정리한다.
        # length/dia 는 raw 좌표 기반이라 유압 결과 불변 — bend 는 표시좌표(ortho)뿐.
        if frozenset((_bkey(a), _bkey(b))) in _branch_keys:
            axd, ayd = _oxy(a); bxd, byd = _oxy(b)
            ddx, ddy = abs(axd - bxd), abs(ayd - byd)
            if ddx > 2.0 and ddy > 2.0:
                if parent_map.get(b) == a:
                    up, dn = (axd, ayd), (bxd, byd)
                elif parent_map.get(a) == b:
                    up, dn = (bxd, byd), (axd, ayd)
                else:
                    up, dn = (axd, ayd), (bxd, byd)
                corner = (dn[0], up[1]) if ddy >= ddx else (up[0], dn[1])
                pipe_dict["bend"] = [int(round(corner[0])), int(round(corner[1]))]
        tables.pipes.append(pipe_dict)
        pipe_label_counter += 1

    # Nozzles
    for i, (h, dist) in enumerate(zip(selection.heads, selection.distances), start=1):
        head_lab = head_node_label[h.pos]
        tables.nozzles.append({
            "label": str(i), "in": head_lab, "out": f"@/{i}",
            "status": "1", "lib": "SP-HEAD",
            "flow_m3s": 0.00133333333, "flow_lmin": 80,
        })

    # ====== Fittings ======
    # 1) 흡수된 elbow → fitting (collinear merge 시 기록된 elbow_fittings 활용)
    for edge_key, elbows in selection.elbow_fittings.items():
        pipe_label = edge_key_to_pipe.get(edge_key)
        if not pipe_label:
            continue
        pipe = next((p for p in tables.pipes if p["label"] == pipe_label), None)
        if not pipe:
            continue
        for _node_pos, angle_deg in elbows:
            # 정확히 45도 근처 (43.5~46.5) 만 elbow-45 — 참조는 elbow-45 1개뿐
            if 43.5 <= angle_deg <= 46.5:
                ftype = "elbow-45"
            elif angle_deg >= 70:
                ftype = "elbow"
            else:
                continue
            tables.fittings.append({
                "pipe": pipe_label, "in": pipe["in"], "out": pipe["out"],
                "type": ftype, "count": "1",
            })
    # 2) 차수 ≥ 3 노드 → tee (in 노드 기준)
    node_degrees: Counter[str] = Counter()
    node_pipes: dict[str, list[dict]] = defaultdict(list)
    for p in tables.pipes:
        node_degrees[p["in"]] += 1
        node_degrees[p["out"]] += 1
        node_pipes[p["in"]].append(p)
        node_pipes[p["out"]].append(p)
    for p in tables.pipes:
        if node_degrees[p["in"]] >= 3:
            tables.fittings.append({
                "pipe": p["label"], "in": p["in"], "out": p["out"],
                "type": "tee", "count": "1",
            })

    # (95도까지 흡수 모드 — preserved elbow 별도 검출 불필요)

    # ====== Equipment ======
    # 1) FX flexible — pipe_entities 에서 'SP 후렉시블' LWPOLYLINE 찾기
    fx_count = 0
    if pipe_entities:
        # 헤드 좌표를 라벨로 매핑 (스냅된 위치)
        head_pos_set = {h.pos for h in selection.heads}
        head_pos_to_label = {h.pos: head_node_label[h.pos] for h in selection.heads}
        for en in pipe_entities:
            if en.get("l") != "SP 후렉시블":
                continue
            if en["t"] != "PL":
                continue
            pts = en["p"]
            if len(pts) < 2:
                continue
            start = _round_pt(pts[0][0], pts[0][1])
            end = _round_pt(pts[-1][0], pts[-1][1])
            # FX 한쪽 endpoint 가 head, 다른 쪽이 subgraph 노드면 그 pipe 에 FX 부착
            head_end = None; pipe_end = None
            for ep in (start, end):
                # 가장 가까운 head 찾기 (within 500mm)
                best_h = None; best_d = float("inf")
                for hp in head_pos_set:
                    d = math.hypot(ep[0] - hp[0], ep[1] - hp[1])
                    if d < best_d:
                        best_d = d; best_h = hp
                if best_h is not None and best_d <= 500.0:
                    head_end = best_h
                else:
                    pipe_end = ep
            if head_end is None:
                continue
            head_label = head_pos_to_label[head_end]
            # 그 head 가 in 노드인 nozzle 의 pipe 를 찾자 — 단순화: head_label 이 in/out 인 첫 pipe
            attached_pipe = next((p for p in tables.pipes if p["in"] == head_label or p["out"] == head_label), None)
            if not attached_pipe:
                continue
            # 중복 방지 — 같은 head 에 이미 FX 가 부착되어 있으면 skip
            already = any(
                eq["desc"] == "FX" and (eq["in"] == head_label or eq["out"] == head_label)
                for eq in tables.equipment
            )
            if already:
                continue
            fx_count += 1
            # FX 등가길이 — 도면의 물리 길이가 아니라 형식승인/제품 스펙 기준의 고정값.
            # 기본 프리셋 "평균"(구 A사 유형, 15.6m) 채택 (FX_SPEC_PROFILES 참조). 도면 물리길이는
            # fx_len_mm 으로 별도 계산해 drawing_len_mm(QA/대사용)에만 기록하고 eq_len 에는 쓰지 않음.
            fx_len_mm = 0.0
            for p0, p1 in zip(pts, pts[1:]):
                fx_len_mm += math.hypot(p1[0] - p0[0], p1[1] - p0[1])
            tables.equipment.append({
                "pipe": attached_pipe["label"], "in": attached_pipe["in"], "out": attached_pipe["out"],
                "label": str(fx_count + 1), "desc": "FX",
                "eq_len": FX_SPEC_PROFILES[FX_DEFAULT_PROFILE]["eq_len_m"],
                "rel_pos": 0.5,
                "spec_ref": FX_DEFAULT_PROFILE,
                "source": "extracted",
                "override_flag": False,
                "override_note": "",
                "drawing_len_mm": round(fx_len_mm, 1),
            })

    # 1.5) FX 보충 — 헤드 30개 모두 FX 1개씩 (참조 패턴: 각 head 에 FX flexible 1개)
    head_with_fx = {
        eq["in"] if eq["in"] in head_node_label.values() else eq["out"]
        for eq in tables.equipment if eq["desc"] == "FX"
    }
    for h, dist in zip(selection.heads, selection.distances):
        head_label = head_node_label[h.pos]
        if head_label in head_with_fx:
            continue
        attached_pipe = next((p for p in tables.pipes if p["in"] == head_label or p["out"] == head_label), None)
        if not attached_pipe:
            continue
        fx_count += 1
        tables.equipment.append({
            "pipe": attached_pipe["label"], "in": attached_pipe["in"], "out": attached_pipe["out"],
            "label": str(fx_count + 1), "desc": "FX",
            "eq_len": FX_SPEC_PROFILES[FX_DEFAULT_PROFILE]["eq_len_m"], "rel_pos": 0.5,
            "spec_ref": FX_DEFAULT_PROFILE,
            "source": "supplemented",
            "override_flag": False,
            "override_note": "",
            "drawing_len_mm": None,
        })

    # 2) 알람밸브 (A/V) — src_label 이 in/out 인 첫 pipe 에 부착
    av_pipe = next((p for p in tables.pipes if p["in"] == src_label or p["out"] == src_label), None)
    if av_pipe:
        tables.equipment.insert(0, {
            "pipe": av_pipe["label"], "in": av_pipe["in"], "out": av_pipe["out"],
            "label": "1", "desc": "A/V",
            "eq_len": AV_EQ_LEN_M, "rel_pos": 0.5,
            "spec_ref": "AV_STD",
            "source": "supplemented",
            "override_flag": False,
            "override_note": "",
            "drawing_len_mm": None,
        })

    # Meta
    tables.meta = [
        ("원본 파일", project_title),
        ("SDF 버전", "1.8  (0)"),
        ("생성 모듈", "Remote 30 프로토타입"),
        ("선정 헤드 수", str(len(selection.heads))),
        ("subgraph 노드 수", str(len(label_to_pos))),
        ("subgraph 파이프 수", str(len(tables.pipes))),
        (f"세대내부 CPVC 배관 (C={CPVC_C_FACTOR})", f"{cpvc_pipe_count} / 강관 {len(tables.pipes) - cpvc_pipe_count}"),
        ("Fittings", str(len(tables.fittings))),
        ("Equipment", str(len(tables.equipment))),
        ("알람밸브 좌표 (snap)", f"({selection.source_pos[0]:.1f}, {selection.source_pos[1]:.1f})"),
        ("source 자동 식별 방식", selection.source_kind),
        ("Diameter 추론 — DXF text 매칭", str(diameter_source_counter.get("text", 0))),
        ("Diameter 추론 — NFPC 별표 1 보강 (text<min)", str(diameter_source_counter.get("nfpc_min", 0))),
        ("Diameter 추론 — NFPC 별표 1 fallback (text 미매칭)", str(diameter_source_counter.get("nfpc_fallback", 0))),
        ("Diameter 텍스트 후보 수 (도면)", str(len(dia_text_pts))),
    ]
    return tables


def write_csv_tables(tables: PipeTables, out_dir: Path, prefix: str) -> dict[str, Path]:
    """5 CSV 출력. 참조 xlsx 의 컬럼 순서와 동일."""
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = {}
    headers = {
        "nodes": ["Label", "Elevation (m)", "I/O node", "Position X (mm)", "Position Y (mm)", "Use spec in scenarios"],
        "pipes": ["Label", "Input node", "Output node", "Type", "Diameter (mm)", "Length (m)", "Elevation (m)",
                  "C-factor", "Status", "Design group", "Fittings", "Equipment"],
        "nozzles": ["Label", "Input node", "Output", "Status", "Library item", "Flow (m³/s)", "Flow (L/min)"],
        "fittings": ["Pipe label", "Input node", "Output node", "Fitting type", "Count"],
        "equipment": ["Pipe label", "Input node", "Output node", "Equipment label", "Description",
                      "Equivalent length (m)", "Rel-position",
                      "Spec profile", "Source", "Override", "Override note", "Drawing length (mm)"],
    }
    rows_map = {
        "nodes": [[n["label"], n["elevation"], n["io_node"], n["x"], n["y"], None] for n in tables.nodes],
        "pipes": [[p["label"], p["in"], p["out"], p["type"], p["dia"], p["length"], p["elev"],
                   p["c"], p["status"], p["group"], None, None] for p in tables.pipes],
        "nozzles": [[n["label"], n["in"], n["out"], n["status"], n["lib"], n["flow_m3s"], n["flow_lmin"]] for n in tables.nozzles],
        "fittings": [[f["pipe"], f["in"], f["out"], f["type"], f["count"]] for f in tables.fittings],
        "equipment": [[e["pipe"], e["in"], e["out"], e["label"], e["desc"], e["eq_len"], e["rel_pos"],
                       e.get("spec_ref", ""), e.get("source", ""), e.get("override_flag", False),
                       e.get("override_note", ""), e.get("drawing_len_mm")] for e in tables.equipment],
    }
    for name in ("nodes", "pipes", "nozzles", "fittings", "equipment"):
        p = out_dir / f"{prefix}_{name}.csv"
        with p.open("w", newline="", encoding="utf-8-sig") as fp:
            w = csv.writer(fp)
            w.writerow(headers[name])
            w.writerows(rows_map[name])
        paths[name] = p
    return paths


def write_xlsx_tables(tables: PipeTables, out_path: Path) -> Path:
    """참조 xlsx 구조 그대로 6 시트 emit."""
    import openpyxl  # local import
    wb = openpyxl.Workbook()
    # default sheet 제거
    wb.remove(wb.active)

    sheet_specs = [
        ("Pipes", ["Label", "Input node", "Output node", "Type", "Diameter (mm)", "Length (m)",
                   "Elevation (m)", "C-factor", "Status", "Design group", "Fittings", "Equipment"],
         [[p["label"], p["in"], p["out"], p["type"], p["dia"], p["length"], p["elev"],
           p["c"], p["status"], p["group"], None, None] for p in tables.pipes]),
        ("Nodes", ["Label", "Elevation (m)", "I/O node", "Position X (mm)", "Position Y (mm)", "Use spec in scenarios"],
         [[n["label"], n["elevation"], n["io_node"], n["x"], n["y"], None] for n in tables.nodes]),
        ("Nozzles", ["Label", "Input node", "Output", "Status", "Library item", "Flow (m³/s)", "Flow (L/min)"],
         [[n["label"], n["in"], n["out"], n["status"], n["lib"], n["flow_m3s"], n["flow_lmin"]] for n in tables.nozzles]),
        ("Fittings", ["Pipe label", "Input node", "Output node", "Fitting type", "Count"],
         [[f["pipe"], f["in"], f["out"], f["type"], f["count"]] for f in tables.fittings]),
        ("Equipment", ["Pipe label", "Input node", "Output node", "Equipment label", "Description",
                       "Equivalent length (m)", "Rel-position",
                       "Spec profile", "Source", "Override", "Override note", "Drawing length (mm)"],
         [[e["pipe"], e["in"], e["out"], e["label"], e["desc"], e["eq_len"], e["rel_pos"],
           e.get("spec_ref", ""), e.get("source", ""), e.get("override_flag", False),
           e.get("override_note", ""), e.get("drawing_len_mm")] for e in tables.equipment]),
        ("Meta", ["항목", "내용"], [[k, v] for k, v in tables.meta]),
    ]
    for name, header, rows in sheet_specs:
        ws = wb.create_sheet(name)
        ws.append(header)
        for r in rows:
            ws.append(r)
    wb.save(out_path)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# 4) Stage 4 — PIPENET SDF emit
# ────────────────────────────────────────────────────────────────────────────


def bundle_result_zip(out_dir: Path, prefix: str) -> Path:
    """결과 폴더의 .sdf + .slf + .xlsx + csv/*.csv 를 .zip 으로 묶음.

    PIPENET 에서 열려면 .sdf 와 .slf 가 동일 폴더에 있어야 하므로,
    사용자에게는 zip 한 번에 받아 unzip 하도록 안내.
    """
    import zipfile
    zip_path = out_dir / f"{prefix}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for suf in (".sdf", ".slf", ".xlsx", ".kfp"):
            p = out_dir / f"{prefix}{suf}"
            if p.is_file():
                zf.write(p, arcname=p.name)
        csv_dir = out_dir / "csv"
        if csv_dir.is_dir():
            for f in sorted(csv_dir.glob(f"{prefix}_*.csv")):
                zf.write(f, arcname=f"csv/{f.name}")
    return zip_path


def write_sdf_tree(tree: "ET.ElementTree", out_path: Path) -> None:
    """ElementTree → PIPENET SDF 직렬화 (DOCTYPE 보존).

    ``ElementTree.write`` 는 ``<!DOCTYPE ...>`` 를 보존하지 못하고 XML 선언도
    ``<?xml version='1.0' encoding='utf-8'?>`` (작은따옴표·소문자) 로 쓴다.
    레퍼런스 (2. 출력 배관망_수작업.sdf) 의 헤더는
    ``<?xml version="1.0" encoding="UTF-8"?>`` + ``<!DOCTYPE Project SYSTEM "spray.dtd">`` —
    DOCTYPE 가 빠지면 일부 PIPENET 설치에서 파일 열기/연산이 거부된다(다른 PC 에서
    연산 오류의 원인). SLF 의 _harden_slf_for_combined 와 동일 패턴으로 헤더를 직접 붙인다.
    """
    body = ET.tostring(tree.getroot(), encoding="unicode")
    Path(out_path).write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE Project SYSTEM "spray.dtd">\n' + body,
        encoding="utf-8",
    )


def _materialize_fx_pipes(tables: "PipeTables") -> tuple["PipeTables", dict, dict]:
    """FX 등가길이 Equipment 를 참조 SDF 구조(실배관 + 등가길이)로 확장.

    변환(각 FX Equipment E, 부모 파이프 P, 헤드노드 H 에 대해):
      1. 새 노드 F 삽입 (좌표=H, 표고=H_elev - FX_RISE_M → F→H 표고차 = FX_RISE_M).
      2. P 의 끝점 중 H 인 것을 F 로 리디렉트 (P: ...→F).
      3. FX 파이프 FXP(F→H) 추가: 호칭경=nominal_dn, 길이=phys_len_m,
         rise=FX_RISE_M, C=c_factor, type=FX_<기하> 스케줄명(SLF Item-name 겸용).
      4. 등가길이 Equipment E 를 FXP 로 이동 (eq_len 불변 — 실배관 위에 얹는 추가 등가길이).
    노즐(in=H)은 그대로 H 에 남는다 → 흐름: ...→F→(FXP)→H→(nozzle)→토출.

    반환: (변환된 tables 복사본,
           {FX파이프 label: 스케줄명},
           {스케줄명: (nominal_dn, inner_dia_mm, c_factor)}).
    호칭경이 같아도 내경이 다르면(예 A사 21.6 vs B사 21.5) 별개 스케줄로 분리된다.
    원본 tables 는 변경하지 않는다(deepcopy).
    """
    import copy as _copy
    t = PipeTables(
        nodes=_copy.deepcopy(tables.nodes),
        pipes=_copy.deepcopy(tables.pipes),
        nozzles=_copy.deepcopy(tables.nozzles),
        fittings=_copy.deepcopy(tables.fittings),
        equipment=_copy.deepcopy(tables.equipment),
        meta=list(tables.meta),
    )
    head_nodes = {str(nz["in"]) for nz in t.nozzles}
    node_by_label = {str(n["label"]): n for n in t.nodes}
    pipe_by_label = {str(p["label"]): p for p in t.pipes}

    def _next_int_label(existing: set) -> int:
        mx = 0
        for lb in existing:
            try:
                v = int(float(lb))
            except (TypeError, ValueError):
                continue
            if v > mx:
                mx = v
        return mx + 1

    node_ctr = _next_int_label(set(node_by_label))
    pipe_ctr = _next_int_label(set(pipe_by_label))

    fx_pipe_sched: dict[str, str] = {}
    fx_geoms: dict[str, tuple] = {}

    for eq in t.equipment:
        if str(eq.get("desc")) != "FX":
            continue
        prof = FX_SPEC_PROFILES.get(str(eq.get("spec_ref"))) or FX_SPEC_PROFILES[FX_DEFAULT_PROFILE]
        p = pipe_by_label.get(str(eq.get("pipe")))
        if p is None:
            continue
        p_in, p_out = str(p["in"]), str(p["out"])
        if p_out in head_nodes:
            head = p_out
        elif p_in in head_nodes:
            head = p_in
        else:
            # 헤드노드 특정 실패 → 변환 생략(종전대로 부모파이프에 등가길이만 — 안전 fallback)
            continue
        h_node = node_by_label.get(head)
        if h_node is None:
            continue

        nominal_dn = int(prof["nominal_dn"])
        inner_dia = float(prof["inner_dia_mm"])
        c_factor = float(prof["c_factor"])
        phys_len = float(prof["phys_len_m"])
        sched = fx_schedule_name(nominal_dn, inner_dia)

        # 1) 새 노드 F
        f_label = str(node_ctr); node_ctr += 1
        f_node = {
            "label": f_label,
            "elevation": float(h_node.get("elevation", 0.0) or 0.0) - FX_RISE_M,
            "io_node": "No",
            "x": h_node["x"], "y": h_node["y"],
        }
        if "display_z" in h_node:
            f_node["display_z"] = h_node["display_z"]
        t.nodes.append(f_node)
        node_by_label[f_label] = f_node

        # 2) 부모 파이프 P 의 head 끝점 → F
        if str(p["out"]) == head:
            p["out"] = f_label
        else:
            p["in"] = f_label

        # 3) FX 파이프 FXP(F→H)
        fxp_label = str(pipe_ctr); pipe_ctr += 1
        fxp = {
            "label": fxp_label,
            "in": f_label, "out": head,
            "type": sched,
            "dia": nominal_dn,          # 호칭경(mm) → bore. 내경은 SLF FX 스케줄 lookup.
            "length": phys_len,
            "elev": FX_RISE_M,
            "c": c_factor,
            "status": "Normal",
            "group": "Unset",
        }
        t.pipes.append(fxp)
        pipe_by_label[fxp_label] = fxp

        # 4) 등가길이 Equipment 를 FXP 로 이동 (eq_len 불변)
        eq["pipe"] = fxp_label
        eq["in"] = f_label
        eq["out"] = head

        fx_pipe_sched[fxp_label] = sched
        fx_geoms[sched] = (nominal_dn, inner_dia, c_factor)

    return t, fx_pipe_sched, fx_geoms


def _rewrite_slf_fx_schedules(slf_path: Path, fx_geoms: dict) -> None:
    """동봉 SLF 의 정적 <FX> 스케줄을 사용된 규격 기하별 FX_<기하> 스케줄로 치환.

    fx_geoms: {스케줄명: (nominal_dn, inner_dia_mm, c_factor)}.
    각 스케줄 = <Item-name>스케줄명, roughness=FX_SCHEDULE_ROUGHNESS,
    Size-definition internal=inner_dia_mm nominal=nominal_dn 1행.
    SLF DOCTYPE(<!DOCTYPE Library SYSTEM "Library.dtd">) 를 직접 붙여 보존(_harden_slf_for_combined 동일 패턴).
    """
    import xml.etree.ElementTree as _ET
    if not slf_path.is_file() or not fx_geoms:
        return
    try:
        tree = _ET.parse(slf_path)
    except _ET.ParseError:
        return
    root = tree.getroot()
    sec = root.find("Schedule-section")
    if sec is None:
        return
    # 정적 FX 스케줄(Item-name == "FX") 위치를 찾아 제거 → 그 자리에 규격별 스케줄 삽입.
    children = list(sec)
    fx_idx = len(children)
    for i, sch in enumerate(children):
        name_el = sch.find("Item-name")
        if name_el is not None and (name_el.text or "").strip() == "FX":
            fx_idx = i
            sec.remove(sch)
            break
    existing = {
        (s.find("Item-name").text or "").strip()
        for s in sec.findall("Schedule") if s.find("Item-name") is not None
    }
    insert_at = fx_idx
    for sched_name, geom in fx_geoms.items():
        nominal_dn, inner_dia = int(geom[0]), float(geom[1])
        if sched_name in existing:
            continue
        sch = _ET.Element("Schedule", {"poisson-ratio": "Unset", "youngs-modulus": "Unset"})
        _ET.SubElement(sch, "Item-name").text = sched_name
        _ET.SubElement(sch, "Description").text = sched_name
        md = _ET.SubElement(sch, "Metric-definition", {"roughness": ("%g" % FX_SCHEDULE_ROUGHNESS)})
        _ET.SubElement(md, "Size-definition", {
            "external": "Unset",
            "internal": ("%g" % inner_dia),
            "nominal": str(nominal_dn),
        })
        sec.insert(insert_at, sch)
        insert_at += 1
        existing.add(sched_name)

    body = _ET.tostring(root, encoding="unicode")
    slf_path.write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE Library SYSTEM "Library.dtd">\n' + body,
        encoding="utf-8",
    )


def emit_sdf(tables: PipeTables, out_path: Path, *, project_title: str = "Remote 30 Prototype") -> Path:
    """PIPENET SDF emit — pipenet_converter.sdf_writer 의 template_path 활용.

    참조 SDF 를 template 으로 사용하면 Network-spray 의 Nodes/Links 만 우리 데이터로
    교체되고 나머지 (Attributes/Libraries/Graphics 의 Display-options/Link-schemes/
    Node-schemes 등 아이소매트릭 표시 메타데이터) 가 모두 보존된다. 결과 SDF 가
    PIPENET 에서 정상적으로 열리며 isometric 도식도 표시됨.

    구조 (참조와 동일):
        <Project version="1.6  (0)">
          <Network-spray>
            <Title>..</Title>
            <Attributes>..</Attributes>  (template)
            <Libraries>..</Libraries>    (template)
            <Nodes><Node label=.. elevation=.. io-node=..><Position x=.. y=../></Node>...</Nodes>
            <Links>
              <Pipe-set>
                <Pipe-type c-factor=.. ..><Name>KSD 3507</Name>..</Pipe-type>
                <Pipe bore="0.025" input=.. label=.. length=.. output=.. rise=.. roughness-or-c=.. status=..>
                  <Fittings><Fitting count="1" type="tee"/></Fittings>
                  <Components><Equipment description="A/V" equivalent-length=.. label=.. rel-position=../></Components>
                  <Waypoints symbol-segment="0"><Position x=.. y=../></Waypoints>
                </Pipe>
              </Pipe-set>
              <Nozzle input=.. label=.. output="@/N" status="1">
                <Flow-define flow=.."/>
                <Library-item>SP-HEAD</Library-item>
              </Nozzle>
            </Links>
          </Network-spray>
          <Graphics>..</Graphics>  (template — Display-options/Schemes/Text-element)
        </Project>
    """
    # pipenet_converter 가 src layout 이라 sys.path 보강
    import sys as _sys
    _pc_src = Path(__file__).parent / "pipenet_converter" / "src"
    if _pc_src.is_dir() and str(_pc_src) not in _sys.path:
        _sys.path.insert(0, str(_pc_src))
    from pipenet_converter.models import (
        Equipment as PnEquipment,
        Fitting as PnFitting,
        Node as PnNode,
        Nozzle as PnNozzle,
        Pipe as PnPipe,
        PipeNetwork,
    )
    from pipenet_converter.sdf_writer import write_sdf as _write_sdf

    # ── FX 실배관 materialize: 등가길이 Equipment → 실배관(FX_<기하>)+등가길이 로 확장.
    # 이후 모든 노드/파이프 처리는 확장된 복사본(tables) 기준. fx_pipe_sched/fx_geoms 는
    # 후처리에서 FX 파이프를 전용 Pipe-set 으로 분리 + SLF FX 스케줄 동적 생성에 사용.
    tables, _fx_pipe_sched, _fx_geoms = _materialize_fx_pipes(tables)

    network = PipeNetwork(title=project_title)

    # ── 좌표 정규화: DXF bbox 중심 → (0,0), 가장 긴 축 → 약 3000 unit (PIPENET 캔버스 fit)
    _xs = [float(n["x"]) for n in tables.nodes]
    _ys = [float(n["y"]) for n in tables.nodes]
    if _xs and _ys:
        _cx = (min(_xs) + max(_xs)) / 2.0
        _cy = (min(_ys) + max(_ys)) / 2.0
        _longest = max(max(_xs) - min(_xs), max(_ys) - min(_ys))
        _scale = (3000.0 / _longest) if _longest > 1e-9 else 1.0
    else:
        _cx = _cy = 0.0
        _scale = 1.0

    def _xform(x: float, y: float) -> tuple[float, float]:
        return ((x - _cx) * _scale, (y - _cy) * _scale)

    # 노드
    for n in tables.nodes:
        nx, ny = _xform(float(n["x"]), float(n["y"]))
        _meta = {"io_node": n["io_node"]}
        # display_z(표시 전용 z) 가 있으면 x,y 와 **동일 _scale** 로 정규화해 Position z 로
        # 전달 → 최종 KFP/HAS 에서 평면과 비례하는 라이저 입상관. elevation(z=)은 수리
        # 실표고 그대로. (단독망/기존 호출은 display_z 미지정 → Position z 미기록.)
        _dz = n.get("display_z")
        if _dz is not None:
            _meta["display_z"] = float(_dz) * _scale
        network.add_node(PnNode(
            node_id=str(n["label"]),
            x=nx, y=ny, z=float(n["elevation"]),
            node_type="input" if n["io_node"] == "Input" else "no",
            metadata=_meta,
        ))

    # 파이프 (fittings/equipment 부착 위해 미리 dict 인덱싱)
    fittings_by_pipe: dict[str, list[PnFitting]] = defaultdict(list)
    for f in tables.fittings:
        fittings_by_pipe[str(f["pipe"])].append(PnFitting(
            fitting_type=str(f["type"]), count=int(f["count"])
        ))
    equip_by_pipe: dict[str, list[PnEquipment]] = defaultdict(list)
    for e in tables.equipment:
        equip_by_pipe[str(e["pipe"])].append(PnEquipment(
            equipment_id=str(e["label"]),
            description=str(e["desc"]),
            equivalent_length_m=float(e["eq_len"]),
            rel_position=float(e["rel_pos"]),
        ))

    # CPVC(단위세대) 로 표기된 배관 label 집합 — 아래 SDF 후처리에서 KSD 3507
    # Pipe-set → CPVC2 Pipe-set 으로 <Pipe> 요소를 이동시킬 때 사용(재질별 schedule 분리).
    cpvc_labels = {str(p["label"]) for p in tables.pipes
                   if str(p.get("type", "")) == CPVC_PIPE_TYPE}

    for p in tables.pipes:
        pid = str(p["label"])
        network.add_pipe(PnPipe(
            pipe_id=pid,
            from_node=str(p["in"]),
            to_node=str(p["out"]),
            diameter_m=float(p["dia"]) / 1000.0,  # mm → m (PIPENET 표준)
            length_m=float(p["length"]),
            rise_m=float(p.get("elev", 0.0) or 0.0),
            c_factor=float(p["c"]),
            status="normal",
            fittings=fittings_by_pipe.get(pid, []),
            equipment=equip_by_pipe.get(pid, []),
            waypoints=[],
        ))

    # 노즐
    for nz in tables.nozzles:
        in_id = str(nz["in"])
        out_id = str(nz["out"])
        # 노즐 토출구(@/N 대기노드)를 <Node> 로 선언 — 참조 SDF 와 동일하게.
        # 미선언 시 PIPENET 이 토출 대기노드를 해석하지 못해 연산이 실패한다.
        if out_id not in network.nodes:
            anchor = network.nodes.get(in_id)
            if anchor is not None:
                nx, ny, nz_elev = anchor.x, anchor.y, anchor.z
            else:
                nx = ny = 0.0
                nz_elev = 0.0
            network.add_node(PnNode(
                node_id=out_id,
                x=nx, y=ny, z=nz_elev,
                node_type="no",
                metadata={"io_node": "No"},
            ))
        network.add_nozzle(PnNozzle(
            nozzle_id=str(nz["label"]),
            input_node=in_id,
            output_node=out_id,
            flow_m3s=float(nz["flow_m3s"]),
            status=int(nz["status"]),
            library_item=str(nz["lib"]),
        ))

    # 참조 SDF 를 template 로 사용 — Graphics 블록 (아이소매트릭 메타) 자동 보존.
    # 경로 해석: 환경변수 REMOTE30_TEMPLATE_SDF → 모듈 디렉토리 fallback. (resolve_template_sdf 참조)
    template = resolve_template_sdf()
    if template is None:
        warnings.warn(
            f"[remote30_prototype.emit_sdf] Template SDF 를 찾을 수 없음. "
            f"결과 SDF 의 Graphics 블록(아이소매트릭 표시 메타·schemes·Display-options) 이 누락됩니다. "
            f"→ 환경변수 REMOTE30_TEMPLATE_SDF 로 절대 경로 지정, 또는 표준 파일 "
            f"'{TEMPLATE_SDF_FILENAME}' 을 모듈 디렉토리 '{_MODULE_DIR}' 에 두세요.",
            RuntimeWarning, stacklevel=2,
        )
    _write_sdf(network, out_path, template_path=template)

    # ── 표준 라이브러리(.slf) 를 결과 폴더에 동봉 — PIPENET 이 호칭경↔내경 매핑 lookup 용.
    # SLF 는 6 schedule (KSD 3507/3562/3576/DP/CPVC2/FX) + SP-HEAD / INDOOR HYDRANT 노즐 + 표준 펌프 정의를 담은
    # 프로젝트 표준 라이브러리. 모든 수리계산 결과물 SDF 가 이 SLF 를 참조하도록 통일.
    # 경로 해석: 환경변수 REMOTE30_STANDARD_SLF → 모듈 디렉토리 fallback. (resolve_standard_slf 참조)
    import shutil as _shutil
    ref_slf = resolve_standard_slf()
    slf_name = out_path.with_suffix(".slf").name  # 예: prototype_<id>.slf
    slf_dst = out_path.parent / slf_name
    if ref_slf is not None and ref_slf.is_file():
        _shutil.copy2(ref_slf, slf_dst)
        # FX 스케줄 동적 재작성 — 정적 SLF 의 단일 <FX> 를 실제 사용된 규격 기하별
        # FX_<기하> 스케줄 N개로 치환(내경=inner_dia_mm, 호칭=nominal_dn).
        if _fx_geoms:
            _rewrite_slf_fx_schedules(slf_dst, _fx_geoms)
    else:
        warnings.warn(
            f"[remote30_prototype.emit_sdf] 표준 SLF 라이브러리를 찾을 수 없음. "
            f"결과 SDF 에 schedule 라이브러리가 동봉되지 않아 PIPENET 에서 호칭경↔내경 lookup 이 실패해 "
            f"diameter 가 'Unset' 으로 표시됩니다. "
            f"→ 환경변수 REMOTE30_STANDARD_SLF 로 절대 경로 지정, 또는 표준 파일 "
            f"'{STANDARD_SLF_FILENAME}' 을 모듈 디렉토리 '{_MODULE_DIR}' 에 두세요.",
            RuntimeWarning, stacklevel=2,
        )

    # ── Template 잔재 정리 + User-lib 재구성 (동봉 SLF 가리키도록)
    import xml.etree.ElementTree as _ET
    _tree = _ET.parse(out_path)
    _root = _tree.getroot()
    for _g in _root.iter("Graphics"):
        for _te in list(_g.findall("Text-element")):
            _g.remove(_te)
    for _libs in _root.iter("Libraries"):
        for _ul in list(_libs.findall("User-lib")):
            _libs.remove(_ul)
        if slf_dst.is_file():
            # 파일명만 — SDF 와 같은 폴더에 SLF 가 있으면 PIPENET 이 자동 로드
            _ul_new = _ET.Element("User-lib", {"file": slf_name})
            _libs.append(_ul_new)
    for _ns in _root.iter("Network-spray"):
        _titles = list(_ns.findall("Title"))
        for _t in _titles[1:]:
            _ns.remove(_t)
        for _nd in list(_ns.findall("Network-description")):
            _ns.remove(_nd)
    # ── 6 schedule Pipe-type 정의 — 권위 SLF (2. Pipenet_hand.slf) 의 Schedule-section 과 정합.
    # 각 항목: (name, c-factor, [(size_m, max_velocity_m_s), ...])
    # Schedule name 은 SLF 의 Item-name 과 동일해야 PIPENET 이 Pipe-type↔Schedule(내경) 을
    # 바인딩한다. 호칭경 set 은 SLF 의 Size-definition.nominal 과 동일, velocity 컨벤션
    # (≤50mm=6, ≥65mm=10) 은 레퍼런스 알람밸브 SDF 의 Pipe-type 정의에서 도출.
    # DP/FX 처럼 단일 호칭경만 정의된 schedule 은 velocity=10 으로 통일.
    _SCHEDULE_DEFS = [
        ("KSD 3507", "120", [
            (0.015, 6), (0.02, 6), (0.025, 6), (0.032, 6), (0.04, 6), (0.05, 6),
            (0.065, 10), (0.08, 10), (0.09, 10), (0.1, 10), (0.125, 10),
            (0.15, 10), (0.2, 10), (0.25, 10), (0.3, 10),
        ]),
        ("KSD 3562", "120", [
            (0.02, 6), (0.025, 6), (0.032, 6), (0.04, 6), (0.05, 6),
            (0.065, 10), (0.08, 10), (0.09, 10), (0.1, 10), (0.125, 10),
            (0.15, 10), (0.2, 10), (0.25, 10), (0.3, 10),
        ]),
        ("KSD 3576", "120", [
            (0.015, 6), (0.02, 6), (0.025, 6), (0.032, 6), (0.04, 6), (0.05, 6),
            (0.065, 10), (0.08, 10), (0.09, 10), (0.1, 10), (0.125, 10),
            (0.15, 10), (0.2, 10), (0.25, 10), (0.3, 10),
        ]),
        ("DP", "120", [(0.025, 10)]),
        ("CPVC2", "150", [
            (0.015, 6), (0.02, 6), (0.025, 6), (0.032, 6), (0.04, 6), (0.05, 6),
            (0.065, 10), (0.08, 10), (0.09, 10), (0.1, 10),
        ]),
        # FX 는 정적 정의 대신 규격 기하별 FX_<기하> Pipe-set 을 아래에서 동적 생성한다.
    ]

    def _make_pipe_type(name: str, c_factor: str, sizes: list) -> "_ET.Element":
        pt = _ET.Element("Pipe-type", {
            "c-factor": c_factor, "criteria": "velocity", "max-velocity": "10",
        })
        _ET.SubElement(pt, "Name").text = name
        _ET.SubElement(pt, "Schedule").text = name
        for _sz, _vel in sizes:
            _ET.SubElement(pt, "Pipe-size", {
                "Lagging-thickness": "0",
                "size": str(_sz),
                "use": "1",
                "velocity": str(_vel),
            })
        return pt

    # ── FX 실배관 전용 Pipe-set 분리 — writer 는 모든 파이프를 단일 Pipe-set 에 담으므로,
    # FX 파이프(호칭경↔내경이 FX_<기하> 스케줄에 바인딩돼야 함)를 기하별 전용 Pipe-set 으로
    # 이동한다. 아래 KSD 3507 삽입 루프보다 먼저 실행 → 남은 main Pipe-set 만 KSD 3507 이 됨.
    if _fx_pipe_sched:
        for _links in _root.iter("Links"):
            _main_ps = None
            for _ps in _links.findall("Pipe-set"):
                if _ps.find("Pipe") is not None:
                    _main_ps = _ps
                    break
            if _main_ps is None:
                break
            _by_sched: dict = {}
            for _pipe in list(_main_ps.findall("Pipe")):
                _sn = _fx_pipe_sched.get(_pipe.get("label"))
                if _sn is None:
                    continue
                _main_ps.remove(_pipe)
                _by_sched.setdefault(_sn, []).append(_pipe)
            _ins = list(_links).index(_main_ps) + 1
            for _sn, _pipes in _by_sched.items():
                _nominal, _inner, _c = _fx_geoms[_sn]
                _fx_ps = _ET.Element("Pipe-set")
                _fx_ps.append(_make_pipe_type(_sn, ("%g" % _c), [(round(int(_nominal) / 1000.0, 6), 10)]))
                for _pipe in _pipes:
                    _fx_ps.append(_pipe)
                _links.insert(_ins, _fx_ps)
                _ins += 1
            break

    # 현재 모든 추론 파이프는 KSD 3507. populated Pipe-set 에는 KSD 3507 Pipe-type 만 삽입한다.
    # 나머지 5 schedule 은 별도 Pipe-set (Pipe-type 만, Pipe 없음) 으로 정의해 PIPENET UI 의 schedule
    # 선택 드롭다운에 노출 — 추후 분류 로직 (task #8) 이 들어오면 해당 schedule Pipe-set 으로 Pipe 이동.
    for _ps in _root.iter("Pipe-set"):
        if _ps.find("Pipe") is None:
            continue  # 빈 Pipe-set placeholder 는 건너뜀
        if _ps.find("Pipe-type") is not None:
            continue
        _ps.insert(0, _make_pipe_type(*_SCHEDULE_DEFS[0]))

    # ── PIPENET-native 패턴 정합: <Links> 구조를
    #   [empty placeholder] + [populated KSD 3507 Pipe-set] + [other-schedule Pipe-sets]
    # 로 재구성. PIPENET 이 SDF 를 읽을 때 첫 Pipe-set 을 "blank/default" 슬롯으로 예약하고
    # 두 번째부터 Schedule 별 Pipe-type 을 바인딩하는 컨벤션 (레퍼런스 3-1/4-1형, 다이소 모든 SDF 에서 확인).
    # placeholder 없으면 우리 Pipe-type 이 blank 슬롯으로 흡수돼 diameter "Unset" 이슈 발생.
    for _links in _root.iter("Links"):
        _populated = None
        for _child in list(_links):
            if _child.tag == "Pipe-set" and _child.find("Pipe") is not None:
                _populated = _child
                break
        if _populated is None:
            continue
        _idx = list(_links).index(_populated)
        # populated Pipe-set 앞에 빈 placeholder Pipe-set 이 없으면 prepend
        if _idx == 0 or list(_links)[_idx - 1].tag != "Pipe-set" or list(_links)[_idx - 1].find("Pipe") is not None:
            _links.insert(_idx, _ET.Element("Pipe-set"))
            _idx += 1
        # populated Pipe-set 뒤로 나머지 5 schedule Pipe-set 을 추가 (이미 있으면 skip)
        _existing_names = set()
        for _ps in _links.iter("Pipe-set"):
            _name_el = _ps.find("Pipe-type/Name")
            if _name_el is not None and _name_el.text:
                _existing_names.add(_name_el.text)
        _insert_at = _idx + 1
        for _name, _cf, _sizes in _SCHEDULE_DEFS[1:]:
            if _name in _existing_names:
                continue
            _new_ps = _ET.Element("Pipe-set")
            _new_ps.append(_make_pipe_type(_name, _cf, _sizes))
            _links.insert(_insert_at, _new_ps)
            _insert_at += 1
        # ── 재질 분리: CPVC(단위세대) 배관을 KSD 3507 Pipe-set 에서 CPVC2 Pipe-set 으로 이동.
        # label 기준 이동(assign 단계 type=CPVC2) — c-factor 문자열 매칭보다 견고. 유압 불변
        # (length/bore/c 는 <Pipe> 속성 그대로 이동, schedule 바인딩만 CPVC2 로 바뀜).
        if cpvc_labels:
            _cpvc_ps = None
            for _ps in _links.iter("Pipe-set"):
                _nm = _ps.find("Pipe-type/Name")
                if _nm is not None and _nm.text == "CPVC2":
                    _cpvc_ps = _ps
                    break
            if _cpvc_ps is not None:
                for _pipe_el in list(_populated.findall("Pipe")):
                    if _pipe_el.get("label") in cpvc_labels:
                        _populated.remove(_pipe_el)
                        _cpvc_ps.append(_pipe_el)
        break

    write_sdf_tree(_tree, out_path)
    return out_path


def emit_kfp(sdf_path: Path, kfp_path: Path, *, coord_scale: float = 1.0,
             display_geometry: bool = False) -> Path:
    """K-Fire Solver .kfp emit — 이미 쓰여진 .sdf 를 그대로 KFP 로 변환.

    ``coord_scale`` — K-solver 표시좌표 배율(노드 크기 조정용, 기본 1.0). 표시
    전용이라 length_m/elevation_m 기반 유압계산엔 영향 없음.

    ``display_geometry`` — 통합망 전용. 미리보기와 동일한 스키매틱 표시좌표
    (라이저=기둥, display_z)로 비율을 미리보기와 일치시킨다. 단독 도면(기본 False)은
    좌표거리==length_m 자가보정(수리 정확) 유지. kfp_sdf_converter.emit_kfp 참조.

    SDF→KFP 는 검증 완료된 ``kfp_sdf_converter`` 경로를 재사용한다 (노즐→노드 folding,
    head/nozzle 키워드 구분, junction 토폴로지 기반 fitting 재분배, 3D length 재계산,
    K-Fire_Solver 표준 라이브러리 동봉). 최종 .sdf 파일을 입력으로 삼으므로 SDF 와 KFP 가
    항상 동일 네트워크를 가리킨다 (좌표 정규화·Graphics 후처리 반영본 기준).

    sig="" / license_tag="TRIAL" — K-solver 가 둘 다 검증하지 않음을 실측 확인 (kfp-format-notes).
    """
    import sys as _sys
    _repo_root = Path(__file__).parent
    if str(_repo_root) not in _sys.path:
        _sys.path.insert(0, str(_repo_root))
    try:
        from kfp_sdf_converter import convert_sdf_to_kfp as _convert
    except ImportError as exc:
        raise RuntimeError(
            f"[remote30_prototype.emit_kfp] kfp_sdf_converter 모듈을 찾을 수 없어 KFP 출력 불가: {exc}. "
            f"→ kfp_sdf_converter.py 를 모듈 루트 '{_repo_root}' 에 두세요."
        ) from exc
    _convert(sdf_path, kfp_path, coord_scale=coord_scale,
             display_geometry=display_geometry)
    return kfp_path


def emit_has(
    sdf_path: Path,
    has_path: Path,
    *,
    isometric: bool = False,
    iso_z_scale: float = 1.0,
) -> Path:
    """HASS(하스) .has emit — 이미 쓰여진 .sdf 를 그대로 HAS 로 변환.

    SDF→HAS 는 ``has_converter`` (parse_sdf → CommonNetwork → emit_has) 경로를 쓴다.
    붙임 샘플(.has)을 스켈레톤으로 로드해 참조 DB/범례/단위계를 그대로 동봉하므로
    HASS 가 받자마자 열 수 있고, 노즐 K 는 nozzleDataTable 에 자동 등록된다.
    RESULT_* 는 0 으로 비워 HASS 가 재계산하게 한다 (우리는 솔버가 아님).

    ``isometric=True`` 면 화면좌표를 등각투영(계통도)으로 베이크 — 통합모듈이
    HASS/solver 화면처럼 보이도록 켠다. 수리계산엔 영향 없음(표시 전용).
    """
    import sys as _sys
    _repo_root = Path(__file__).parent
    if str(_repo_root) not in _sys.path:
        _sys.path.insert(0, str(_repo_root))
    try:
        from has_converter import convert_sdf_to_has as _convert
    except ImportError as exc:
        raise RuntimeError(
            f"[remote30_prototype.emit_has] has_converter 모듈을 찾을 수 없어 HAS 출력 불가: {exc}. "
            f"→ has_converter.py 를 모듈 루트 '{_repo_root}' 에 두세요."
        ) from exc
    _convert(sdf_path, has_path, isometric=isometric, iso_z_scale=iso_z_scale)
    return has_path


# ────────────────────────────────────────────────────────────────────────────
# Orchestrator — generator yielding progress events
# ────────────────────────────────────────────────────────────────────────────


def run_stages_0_2(
    dxf_path: Path,
    job_id: str,
    alarm_xy: tuple[float, float] | None = None,
    branch_zones: list[tuple[float, float, float, float]] | None = None,
) -> Iterator[dict]:
    """Stage 0~2 만 실행 — 파싱 / 배관망 / 헤드 인식. 결과를 마지막 이벤트로 yield.

    호출자(서버)는 마지막 'stage2_complete' 이벤트의 데이터(detected_heads / pipe_ents / layer_categories /
    bundle.entities/layers/bbox) 를 job state 에 저장해두고, 사용자가 헤드 편집 후 finalize 호출 시
    run_stages_3_5() 에 전달한다.
    """
    t0 = time.time()
    def evt(d):
        d.setdefault("elapsed_ms", int((time.time() - t0) * 1000))
        return d

    # Stage 0: 파싱
    yield evt({"type": "stage", "stage": 0, "status": "running", "label": "DXF 파싱"})
    bundle = parse_dxf_bundle_cached(dxf_path)
    layer_categories = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    yield evt({"type": "entities", "stage": 0,
               "entities": bundle.entities,
               "bbox": {"x_min": bundle.bbox[0], "y_min": bundle.bbox[1],
                        "x_max": bundle.bbox[2], "y_max": bundle.bbox[3]},
               "layers": bundle.layers,
               "summary": {"entity_count": len(bundle.entities),
                           "layer_count": len(bundle.layers),
                           "bbox_diagnostics": bundle.bbox_diagnostics}})
    _diag = bundle.bbox_diagnostics or {}
    _ratio = _diag.get("bbox_ratio", 1.0)
    _outliers = _diag.get("outlier_points", 0)
    _diag_msg = ""
    if _ratio >= 2.0:
        _diag_msg = f" · outlier {_outliers}점 제외 (raw bbox 가 robust 의 {_ratio}× — 자동 보정)"
    yield evt({"type": "stage", "stage": 0, "status": "done",
               "label": f"DXF 파싱 완료 — {len(bundle.entities):,} entity / {len(bundle.layers)} 레이어{_diag_msg}"})

    # Stage 1
    yield evt({"type": "stage", "stage": 1, "status": "running", "label": "건축/기타 레이어 제거 (배관망만)"})
    pipe_ents = filter_pipenet_only(bundle)
    yield evt({"type": "entities", "stage": 1, "entities": pipe_ents,
               "summary": {"entity_count": len(pipe_ents)}})
    yield evt({"type": "stage", "stage": 1, "status": "done",
               "label": f"배관망 추출 완료 — {len(pipe_ents):,} entity"})

    # Stage 2: 헤드 인식
    yield evt({"type": "stage", "stage": 2, "status": "running",
               "label": "도면 내 전체 헤드 후보 인식 (block pattern + CIRCLE/HATCH 시그니처 + 클러스터링)"})
    head_detections = detect_heads(pipe_ents, layer_categories)
    bbox_ents = [{"t": "B", "l": "_head_bbox", "p": list(h.bbox),
                  "k": h.kind, "c": round(h.confidence, 2), "n": h.block_name,
                  "i": idx, "pos": list(h.pos)}
                 for idx, h in enumerate(head_detections)]
    from collections import Counter as _C
    kind_counter: _C = _C()
    for h in head_detections:
        primary = h.kind.split(":")[0] if ":" in h.kind else h.kind
        kind_counter[primary] += 1
    yield evt({"type": "entities", "stage": 2, "entities": bbox_ents,
               "summary": {
                   "head_count": len(head_detections),
                   "by_kind": dict(kind_counter),
                   "avg_confidence": round(sum(h.confidence for h in head_detections) / len(head_detections), 3) if head_detections else 0,
               }})
    yield evt({"type": "stage", "stage": 2, "status": "done",
               "label": f"전체 헤드 {len(head_detections)}개 인식 완료"})

    # ===== Stage 3 (신규): 전체 배관망 그래프 시각화 =====
    # select_worst30_heads 가 사용할 정확한 내부 그래프를 미리 보여줌.
    # epsilon-cluster + 컴포넌트 brigde + 헤드 drop line 모두 포함된 최종 그래프.
    # 그래프 노드 좌표는 raw (DXF 원본) — 격자 정렬 안 됨, 시각화 시 비뚤어짐 없음.
    yield evt({"type": "stage", "stage": 3, "status": "running",
               "label": "전체 배관망 그래프 인식 (epsilon-cluster + 컴포넌트 bridge + 헤드 drop line)"})
    # 모든 좌표(파이프 endpoint, 헤드 INSERT, 알람밸브) 가 같은 NodeIndex 공간을
    # 공유 → 헤드/AV 좌표가 그래프 노드와 정확히 매칭, 별도 nearest fallback 불필요.
    node_index = _NodeIndex()
    graph, edge_len = _build_graph(pipe_ents, node_index=node_index,
                                   layer_categories=layer_categories)
    # 평행 ladder collapse — 관 두 줄 표현 → midline 1줄로. bridge 전에 적용해
    # ladder 양 끝 cap 이 가짜 component 분리 만들지 않도록.
    ladders_collapsed = collapse_parallel_ladders(graph, edge_len)
    # weld_edges: 방향-인지 끝점 용접 (끊긴 직선 배관을 축대로 복원 — 추정 연결)
    # bridge_edges: _bridge_components 가 강제로 이은 연결 (실제 배관 아님)
    # head_drop_edges: 헤드 INSERT 좌표 ↔ 배관 nearest 노드 직선 (실제 배관 아님)
    # 세 종류 모두 "알고리즘이 추정한 연결"이라 시각적으로 구분 렌더.
    # 용접은 bridge 이전에 실행 — 파편을 원래 축대로 먼저 이어 뱀 경로 방지
    # (select_worst30_heads 내부 순서와 동일해 시각화=산출물 토폴로지 일치).
    # 추정연결 허용치는 도면 스케일 비례(적응형) — calc 경로(select_worst30_heads)와
    # 동일 공식이라 시각화=산출물 토폴로지 유지.
    _diag = _graph_diag(graph)
    weld_edges: set = set()
    _weld_dangling_endpoints(graph, edge_len,
                             weld_tol=_adaptive_weld_tol(_diag),
                             weld_cone_deg=_WELD_CONE_DEG,
                             weld_edges_out=weld_edges)
    bridge_edges: set = set()
    for tol in _adaptive_bridge_tols(_diag):
        _bridge_components(graph, edge_len, max_bridge_mm=tol, bridge_edges_out=bridge_edges)
    # 주의: SPT 가 cycle edge 를 제거하면서 bridge_edges 일부도 같이 제거될 수 있음.
    # SPT 적용 후 살아남은 bridge_edges 만 유효.
    # 헤드 drop line — 헤드 INSERT 좌표를 같은 NodeIndex 로 canonicalize
    # (그래프에 이미 같은 epsilon 안 노드 있으면 그 raw 좌표 반환 → drop line 불필요).
    head_drop_edges: set = set()
    head_pos_list = []
    for h in head_detections:  # Stage 2 결과 재사용 (detect_heads 재호출 제거)
        head_pos_list.append(node_index.canonical(h.pos[0], h.pos[1]))
    for hp in head_pos_list:
        if hp in graph:
            # 헤드가 epsilon 안에서 이미 그래프 노드와 일치 → drop line 불필요
            continue
        nearest = _nearest_graph_node(graph, hp)
        if nearest is None or hp == nearest:
            continue
        d = math.hypot(hp[0] - nearest[0], hp[1] - nearest[1])
        if d > 1e-3 and d <= HEAD_BRIDGE_MAX_MM:
            graph.setdefault(hp, set()).add(nearest)
            graph[nearest].add(hp)
            key = (min(hp, nearest), max(hp, nearest))
            edge_len[key] = d
            head_drop_edges.add(key)

    # ── Spanning Tree 강제 — 가지식 트리화 (cycle 제거)
    # AV-rooted Dijkstra SPT. 도달 가능한 노드는 AV 까지 최단 경로 트리, 다른
    # component 는 각자 임의 root. 트리 외 edge 는 graph 에서 제거 + removed
    # set 으로 회수 → 시각화에서 별도 카테고리로 표시 가능.
    # source 가 AV 좌표 (NodeIndex canonicalized) — 그래프 노드와 정확 매칭.
    spt_source = None
    if alarm_xy is not None:
        spt_source = node_index.canonical(float(alarm_xy[0]), float(alarm_xy[1]))
        if spt_source not in graph:
            spt_source = _nearest_graph_node(graph, spt_source)
    # 분기영역 지정인데 수동 알람밸브가 없으면 — calc 경로(select_worst30_heads)처럼
    # source 를 자동결정해 그 노드에 루팅해야 영역 밖 corridor 제한이 걸린다.
    # (안 하면 spt_source=None → _restrict_to_branch_region 이 no-op 으로 빠져
    #  영역 밖 주배관 꼬임이 그대로 남는다.) branch_zones 미지정 시엔 손대지 않아
    # 기존 렌더/골든 불변.
    if spt_source is None and branch_zones:
        # 자동 검출된 알람밸브가 있으면 그 노드에 루팅. (highest-degree 등 임의
        # 노드 폴백은 금지 — corridor 가 엉뚱한 곳에서 출발해 영역에 도달 못하면
        # no-op 이 되거나 잘못된 주배관을 남긴다. source 를 모르면 제한하지 않고
        # 그대로 두어, 프론트가 "알람밸브를 먼저 지정" 하도록 안내한다.)
        _auto_pt, _ = _find_source(pipe_ents, layer_categories)
        if _auto_pt is not None:
            _auto_pt = node_index.canonical(_auto_pt[0], _auto_pt[1])
            if _auto_pt not in graph and graph:
                _auto_pt = _nearest_graph_node(graph, _auto_pt)
            if _auto_pt in graph:
                spt_source = _auto_pt
    # 분기영역 지정 시 — 영역 밖을 source→영역 단일 corridor 로 제한(주배관 하나).
    # SPT 이전에 적용해 시각화=산출물 토폴로지 일치. 미지정 시 no-op(불변).
    # 추정연결(weld∪bridge∪drop) — SPT/corridor 라우팅에 penalty 부여해 실배관 우선.
    # calc 경로(select_worst30_heads)와 동일 원칙이라 시각화=산출물 토폴로지 유지.
    _penalty_keys = weld_edges | bridge_edges | head_drop_edges
    region_nodes = _restrict_to_branch_region(graph, edge_len, spt_source, branch_zones,
                                              penalty_keys=_penalty_keys)
    tree_edges_set, removed_cycle_edges = force_spanning_tree(
        graph, edge_len, source=spt_source, penalty_keys=_penalty_keys)
    # SPT 가 일부 weld/bridge/drop edge 도 제거할 수 있음 — 살아남은 것만 유효
    weld_edges &= tree_edges_set
    bridge_edges &= tree_edges_set
    head_drop_edges &= tree_edges_set

    # edge entity emit — 4종 구분 (실배관 / bridge / drop / 제거된 cycle)
    graph_ents = []
    seen_edges: set = set()
    weld_emitted = 0
    bridge_emitted = 0
    head_drop_emitted = 0
    cycle_emitted = 0
    # 1) 트리 edge (실제 토폴로지) — graph 에 남아 있음
    for u, neighbors in graph.items():
        for v in neighbors:
            key = (min(u, v), max(u, v))
            if key in seen_edges:
                continue
            seen_edges.add(key)
            if key in weld_edges:
                layer = "_graph_weld"
                weld_emitted += 1
            elif key in bridge_edges:
                layer = "_graph_bridge"
                bridge_emitted += 1
            elif key in head_drop_edges:
                layer = "_graph_head_drop"
                head_drop_emitted += 1
            else:
                layer = "_graph_edge"
            graph_ents.append({"t": "L", "l": layer, "p": [u[0], u[1], v[0], v[1]]})
    # 2) 제거된 cycle edge — 별도 카테고리 (회색 매우 흐릿, 참고용).
    #    단, 분기영역 안의 루프는 옵션 B(표시만 보존): 계산은 트리라 SPT 가 제거했지만
    #    실선(_graph_loop)으로 보여 실제 배관임을 표현. 영역 밖/미지정은 종전대로 흐릿.
    loop_emitted = 0
    for (u, v) in removed_cycle_edges:
        if region_nodes and u in region_nodes and v in region_nodes:
            graph_ents.append({"t": "L", "l": "_graph_loop", "p": [u[0], u[1], v[0], v[1]]})
            loop_emitted += 1
        else:
            graph_ents.append({"t": "L", "l": "_graph_removed_cycle", "p": [u[0], u[1], v[0], v[1]]})
            cycle_emitted += 1
    # junction 노드 (차수 ≥ 3) 만 점으로
    junction_count = 0
    for n, neighbors in graph.items():
        if len(set(neighbors)) >= 3:
            graph_ents.append({"t": "C", "l": "_graph_junction", "c": [n[0], n[1]], "r": 80.0})
            junction_count += 1

    # ── 알람밸브(source) 시각화 — 사용자 지정 좌표 또는 자동 식별
    # NodeIndex 로 canonicalize → 그래프 노드와 epsilon 안에서 일치 가능.
    if alarm_xy is not None:
        src_raw_pt = node_index.canonical(float(alarm_xy[0]), float(alarm_xy[1]))
        src_kind_preview = "manual"
    else:
        src_raw_pt, src_kind_preview = _find_source(pipe_ents, layer_categories)
        if src_raw_pt is not None:
            src_raw_pt = node_index.canonical(src_raw_pt[0], src_raw_pt[1])
    src_bridge_preview = 0.0
    src_far = False
    if src_raw_pt is not None and graph:
        src_nearest_pt = _nearest_graph_node(graph, src_raw_pt)
        if src_nearest_pt is not None:
            src_bridge_preview = math.hypot(src_raw_pt[0] - src_nearest_pt[0],
                                            src_raw_pt[1] - src_nearest_pt[1])
            src_far = src_bridge_preview > SOURCE_BRIDGE_MAX_MM
            # source 점 + nearest 점 + drop-line
            graph_ents.append({"t": "C", "l": "_alarm_source",
                               "c": [src_raw_pt[0], src_raw_pt[1]], "r": 150.0})
            if src_bridge_preview > 1e-3:
                graph_ents.append({
                    "t": "L", "l": "_alarm_drop_line",
                    "p": [src_raw_pt[0], src_raw_pt[1], src_nearest_pt[0], src_nearest_pt[1]],
                })
                graph_ents.append({"t": "C", "l": "_alarm_attach",
                                   "c": [src_nearest_pt[0], src_nearest_pt[1]], "r": 90.0})

    real_edge_count = len(seen_edges) - weld_emitted - bridge_emitted - head_drop_emitted
    summary = {
        "node_count": len(graph),
        "edge_count": len(seen_edges),
        "real_edge_count": real_edge_count,
        "weld_edge_count": weld_emitted,
        "bridge_edge_count": bridge_emitted,
        "head_drop_edge_count": head_drop_emitted,
        "junction_count": junction_count,
        "components": len(_connected_components(graph)),
        "ladders_collapsed": ladders_collapsed,
        "removed_cycle_edges": cycle_emitted,
        "loop_edge_count": loop_emitted,
        "source_pos": list(src_raw_pt) if src_raw_pt else None,
        "source_kind": src_kind_preview if src_raw_pt else "none",
        "source_bridge_dist_mm": round(src_bridge_preview, 1),
        "source_far_from_pipes": src_far,
    }
    yield evt({"type": "entities", "stage": 3, "entities": graph_ents, "summary": summary})
    label = (
        f"가지식 트리 — {len(graph)} 노드 / 실배관 {real_edge_count} edge"
        f" / 용접 {weld_emitted} / bridge {bridge_emitted} / 헤드 drop {head_drop_emitted} / 분기 {junction_count}개"
        f" / ladder 합성 {ladders_collapsed} / cycle 제거 {cycle_emitted}"
        + (f" / 루프 {loop_emitted}" if loop_emitted else "")
    )
    if src_raw_pt is not None:
        label += f" · 알람밸브 ↔ 배관망 {src_bridge_preview:.0f}mm"
        if src_far:
            label += " ⚠너무 멈"
    yield evt({"type": "stage", "stage": 3, "status": "done", "label": label})

    # 헤드 편집 일시정지 — 다음은 stage 4~6 (select30 / tables / SDF) 가 run_stages_3_5() 처리
    yield evt({"type": "awaiting_finalize",
               "head_count": len(head_detections),
               "pause_message": "Stage 3 완료. 헤드 객체 수정 후 [배관망 완성] 클릭 시 Stage 4~6 진행."})


def run_stages_3_5(
    dxf_path: Path,
    out_dir: Path,
    job_id: str,
    pipe_ents: list[dict],
    layer_categories: dict[str, str],
    detected_heads_pos: list[tuple[float, float]],
    *,
    k_heads: int = 30,
    alarm_xy: tuple[float, float] | None = None,
    user_added_heads: list[tuple[float, float]] | None = None,
    user_deleted_indices: list[int] | None = None,
    zones: list[tuple[float, float, float, float]] | None = None,
    branch_zones: list[tuple[float, float, float, float]] | None = None,
) -> Iterator[dict]:
    """사용자 편집 결과를 받아 Stage 3~5 실행.

    edited_heads = detected_heads - deleted_indices + user_added
    그 다음 select_worst30_heads(zones=zones, manual_heads=edited_heads).
    """
    t0 = time.time()
    def evt(d):
        d.setdefault("elapsed_ms", int((time.time() - t0) * 1000))
        return d

    # 편집된 헤드 목록 구성
    deleted = set(user_deleted_indices or [])
    edited_heads = [pos for i, pos in enumerate(detected_heads_pos) if i not in deleted]
    if user_added_heads:
        edited_heads.extend(user_added_heads)

    # Stage 4 (기존 3 에서 시프트)
    src_label = "수동 좌표" if alarm_xy else "자동 식별"
    zone_info = f"영역 {len(zones)}개" if zones else "전체"
    yield evt({"type": "stage", "stage": 4, "status": "running",
               "label": f"가장 불리한 {k_heads} 헤드 선정 (알람밸브 {src_label}, {zone_info}, 편집 후 {len(edited_heads)} 헤드 후보)"})
    # 대용량 도면에서 select_worst30_heads 는 수십 초 걸릴 수 있어, 워커 스레드에서
    # 실행하고 progress_cb → 큐 → substep SSE 로 진행바를 채운다. 계산 로직은
    # 콜백만 받을 뿐 그대로라 산출물은 불변(골든 통과).
    import queue as _queue, threading as _threading
    _pq: _queue.Queue = _queue.Queue()
    _box: dict = {}

    def _sel_cb(frac, msg):
        _pq.put((float(frac), str(msg)))

    def _sel_run():
        try:
            _box["r"] = select_worst30_heads(
                pipe_ents, layer_categories, k=k_heads, manual_source=alarm_xy,
                manual_heads=edited_heads, zones=zones, branch_zones=branch_zones,
                progress_cb=_sel_cb)
        except BaseException as _e:  # noqa: BLE001 — 워커 예외를 메인으로 전달
            _box["e"] = _e
        finally:
            _pq.put(None)

    _th = _threading.Thread(target=_sel_run, name="select_worst30", daemon=True)
    _th.start()
    while True:
        _item = _pq.get()
        if _item is None:
            break
        _frac, _msg = _item
        yield evt({"type": "substep", "stage": 4,
                   "progress": round(_frac, 3), "label": _msg})
    _th.join()
    if "e" in _box:
        raise _box["e"]
    selection = _box["r"]
    subgraph_ents = []
    _sg_ortho = orthogonalize_edge_positions(
        selection.edges,
        head_points=[h.pos for h in selection.heads],
        source_point=selection.source_pos)

    def _sg_xy(p):
        return _sg_ortho.get((round(float(p[0]), 3), round(float(p[1]), 3)),
                             (float(p[0]), float(p[1])))
    for a, b, _len in selection.edges:
        pa, pb = _sg_xy(a), _sg_xy(b)
        subgraph_ents.append({"t": "L", "l": "_subgraph", "p": [pa[0], pa[1], pb[0], pb[1]]})
    for h in selection.heads:
        subgraph_ents.append({"t": "C", "l": "_subgraph_head", "c": list(_sg_xy(h.pos)), "r": 80.0})
    if selection.source_pos is not None:
        subgraph_ents.append({"t": "C", "l": "_alarm_valve", "c": list(_sg_xy(selection.source_pos)), "r": 150.0})
    yield evt({"type": "entities", "stage": 4, "entities": subgraph_ents,
               "summary": {
                   "selected_heads": len(selection.heads),
                   "subgraph_edges": len(selection.edges),
                   "subgraph_nodes": len(selection.nodes_in_subgraph),
                   "max_distance_m": round(max(selection.distances) / 1000.0, 2) if selection.distances else 0,
                   "source_kind": selection.source_kind,
                   "source_pos": list(selection.source_pos) if selection.source_pos else None,
               }})
    yield evt({"type": "stage", "stage": 4, "status": "done",
               "label": f"선정 완료 — 헤드 {len(selection.heads)}개 / 경로 {len(selection.edges)} edge"})

    # Stage 5: 5 테이블 (기존 4)
    yield evt({"type": "stage", "stage": 5, "status": "running", "label": "Nodes/Pipes/Nozzles/Fittings/Equipment 테이블 생성"})
    tables = build_input_tables(selection, pipe_entities=pipe_ents, project_title=dxf_path.stem,
                                cpvc_zones=zones)
    csv_dir = out_dir / "csv"
    csv_paths = write_csv_tables(tables, csv_dir, prefix=f"prototype_{job_id}")
    xlsx_path = out_dir / f"prototype_{job_id}.xlsx"
    write_xlsx_tables(tables, xlsx_path)
    yield evt({"type": "tables_preview", "stage": 5,
               "tables": {
                   "nodes": tables.nodes[:8], "pipes": tables.pipes[:8],
                   "nozzles": tables.nozzles[:8], "fittings": tables.fittings[:8],
                   "equipment": tables.equipment[:8], "meta": tables.meta,
               },
               "counts": {
                   "nodes": len(tables.nodes), "pipes": len(tables.pipes),
                   "nozzles": len(tables.nozzles), "fittings": len(tables.fittings),
                   "equipment": len(tables.equipment),
               }})
    yield evt({"type": "stage", "stage": 5, "status": "done",
               "label": f"5 테이블 생성 완료 — Pipes {len(tables.pipes)} / Nodes {len(tables.nodes)} / Nozzles {len(tables.nozzles)}"})

    # 신축배관(FX) 검토 게이트 — stage 6(emit) 는 run_stage_6_emit 으로 분리했다.
    # 헤드 편집 게이트(stage2_complete)와 동일한 패턴: 서버가 이 payload 를 job state 에
    # 저장 → 웹 편집기(신축배관 검토 패널) → POST .../fx/finalize → run_stage_6_emit.
    yield evt({"type": "stage5_complete",
               "tables": tables.as_dict(),
               "project_title": dxf_path.stem,
               "fx_review": {
                   "equipment": tables.equipment,   # 전량 — [:8] 캡 금지
                   "profiles": FX_SPEC_PROFILES,     # 편집기 드롭다운용
                   "default_profile": FX_DEFAULT_PROFILE,
               }})


def _validate_edited_equipment(edited: list[dict], original: list[dict]) -> tuple[list[dict], list[dict]]:
    """웹 편집 결과를 검증·정규화하여 (equipment, warnings) 반환.

    규칙(prompt Task 3):
      - eq_len 은 float > 0 (아니면 ValueError → 확정 차단)
      - spec_ref 는 FX_SPEC_PROFILES 에 존재하거나 override_flag=True 여야 함
      - 사용자가 값을 바꾼 행은 source="manual", override_flag=True 강제
      - 프로파일 기준 ±50% 초과 편차는 warning 이벤트로 통지하되 차단 안 함
    original 은 label 기준으로 매칭해 '변경 여부'를 판정한다.
    """
    orig_by_label = {e.get("label"): e for e in original}
    out_rows: list[dict] = []
    warns: list[dict] = []
    for row in edited:
        r = dict(row)
        label = r.get("label")
        # eq_len 검증
        try:
            eq = float(r.get("eq_len"))
        except (TypeError, ValueError):
            raise ValueError(f"FX 등가길이(eq_len)가 숫자가 아님: {label!r} → {r.get('eq_len')!r}")
        if not (eq > 0):
            raise ValueError(f"FX 등가길이(eq_len)는 0보다 커야 함: {label!r} → {eq}")
        r["eq_len"] = eq

        spec_ref = r.get("spec_ref")
        base = orig_by_label.get(label)
        # 값이 원본과 달라졌는지 판정 (eq_len 또는 spec_ref 변경)
        changed = False
        if base is not None:
            try:
                base_eq = float(base.get("eq_len"))
            except (TypeError, ValueError):
                base_eq = None
            if base_eq is None or abs(base_eq - eq) > 1e-6:
                changed = True
            if base.get("spec_ref") != spec_ref:
                changed = True
        else:
            changed = True  # 원본에 없던 행(수동 추가)

        if changed:
            r["source"] = "manual"
            r["override_flag"] = True
        else:
            r.setdefault("source", base.get("source") if base else "supplemented")
            r.setdefault("override_flag", bool(r.get("override_flag", False)))

        # spec_ref 유효성 — 알려진 프로파일이 아니면 override_flag 필수
        if spec_ref not in FX_SPEC_PROFILES and spec_ref != "AV_STD":
            if not r.get("override_flag"):
                raise ValueError(
                    f"규격(spec_ref)이 미등록({spec_ref!r})인데 override_flag 가 아님: {label!r}. "
                    f"직접 입력 시 override 로 표시해야 함.")

        # ±50% 편차 경고 (프로파일 기준값이 있을 때만)
        prof = FX_SPEC_PROFILES.get(spec_ref)
        if prof is not None:
            ref_eq = prof["eq_len_m"]
            if ref_eq > 0 and abs(eq - ref_eq) / ref_eq > 0.5:
                warns.append({
                    "type": "warning", "label": label, "spec_ref": spec_ref,
                    "eq_len": eq, "ref_eq_len": ref_eq,
                    "message": f"{label}: 등가길이 {eq}m 가 규격({spec_ref}) 기준 {ref_eq}m 대비 ±50% 초과. 확인 요망.",
                })
        out_rows.append(r)
    return out_rows, warns


def run_stage_6_emit(
    out_dir: Path,
    job_id: str,
    tables: PipeTables,
    edited_equipment: list[dict] | None = None,
    *,
    project_title: str = "Remote 30 Prototype",
) -> Iterator[dict]:
    """Stage 6 (SDF emit + SLF 동봉 + KFP + zip) — 신축배관 검토 게이트 이후 단계.

    edited_equipment 가 오면 검증(_validate_edited_equipment) 후 tables.equipment 를 교체.
    None 이면 원본 그대로 emit (편집 없이 확정 = 회귀 없음).
    """
    t0 = time.time()
    def evt(d):
        d.setdefault("elapsed_ms", int((time.time() - t0) * 1000))
        return d

    if edited_equipment is not None:
        try:
            new_equipment, warns = _validate_edited_equipment(edited_equipment, tables.equipment)
        except ValueError as _ve:
            yield evt({"type": "error", "stage": 6, "message": str(_ve)})
            return
        for w in warns:
            yield evt(dict(w))
        tables.equipment = new_equipment

    yield evt({"type": "stage", "stage": 6, "status": "running", "label": "PIPENET SDF emit + .slf 동봉 + zip 묶음"})
    sdf_path = out_dir / f"prototype_{job_id}.sdf"
    emit_sdf(tables, sdf_path, project_title=project_title)
    slf_path = sdf_path.with_suffix(".slf")
    # KFP (K-Fire Solver) — 최종 SDF 를 그대로 변환. 실패해도 SDF 출력은 유지.
    kfp_path = out_dir / f"prototype_{job_id}.kfp"
    kfp_ok = False
    try:
        emit_kfp(sdf_path, kfp_path)
        kfp_ok = kfp_path.is_file()
    except Exception as _kfp_exc:  # noqa: BLE001 — KFP 실패가 전체 파이프라인을 막지 않도록
        warnings.warn(f"[remote30_prototype] KFP emit 실패 (SDF 는 정상): {_kfp_exc}", RuntimeWarning, stacklevel=2)
    zip_path = bundle_result_zip(out_dir, prefix=f"prototype_{job_id}")
    _kfp_label = f" + KFP {kfp_path.stat().st_size/1024:.1f}KB" if kfp_ok else ""
    # SLF 는 표준 라이브러리 자산(resolve_standard_slf)이 있어야 동봉됨 — 없으면 emit_sdf 가
    # 경고만 내고 .slf 를 만들지 않으므로 stat() 전에 존재 확인(WinError 2 방지).
    _slf_label = f" + SLF {slf_path.stat().st_size/1024:.1f}KB" if slf_path.is_file() else " (SLF 자산 없음 — 미동봉)"
    yield evt({"type": "stage", "stage": 6, "status": "done",
               "label": f"SDF {sdf_path.stat().st_size/1024:.1f}KB{_slf_label}{_kfp_label} + ZIP {zip_path.stat().st_size/1024:.1f}KB"})

    outputs = {
        "sdf": sdf_path.name,
    }
    xlsx_path = out_dir / f"prototype_{job_id}.xlsx"
    if xlsx_path.is_file():
        outputs["xlsx"] = xlsx_path.name
    csv_dir = out_dir / "csv"
    for _k in ("nodes", "pipes", "nozzles", "fittings", "equipment"):
        _cp = csv_dir / f"prototype_{job_id}_{_k}.csv"
        if _cp.is_file():
            outputs[f"csv_{_k}"] = _cp.name
    if slf_path.is_file():
        outputs["slf"] = slf_path.name
    if kfp_ok:
        outputs["kfp"] = kfp_path.name
    if zip_path.is_file():
        outputs["zip"] = zip_path.name
    yield evt({"type": "done", "outputs": outputs, "out_dir": str(out_dir)})


def run_stages_3_6(
    dxf_path: Path,
    out_dir: Path,
    job_id: str,
    pipe_ents: list[dict],
    layer_categories: dict[str, str],
    detected_heads_pos: list[tuple[float, float]],
    *,
    k_heads: int = 30,
    alarm_xy: tuple[float, float] | None = None,
    user_added_heads: list[tuple[float, float]] | None = None,
    user_deleted_indices: list[int] | None = None,
    zones: list[tuple[float, float, float, float]] | None = None,
) -> Iterator[dict]:
    """하위호환 원샷 래퍼 — 기존 run_stages_3_5(스테이지 3~6 일괄) 동작 재현.

    stage5_complete 게이트 없이 테이블 생성 후 곧바로 emit 까지 실행한다. FX 편집을
    거치지 않는 자동/배치 경로가 있으면 이 함수를 쓴다. 웹 게이트 경로는
    run_stages_3_5 → (편집) → run_stage_6_emit 을 쓴다.
    """
    tables: PipeTables | None = None
    for ev in run_stages_3_5(
        dxf_path, out_dir, job_id, pipe_ents, layer_categories, detected_heads_pos,
        k_heads=k_heads, alarm_xy=alarm_xy, user_added_heads=user_added_heads,
        user_deleted_indices=user_deleted_indices, zones=zones,
    ):
        if ev.get("type") == "stage5_complete":
            tables = PipeTables.from_dict(ev["tables"])
            continue  # 게이트 이벤트는 원샷 경로에서 소비만 하고 재전달하지 않음
        yield ev
    if tables is None:
        return
    yield from run_stage_6_emit(out_dir, job_id, tables,
                                edited_equipment=None, project_title=dxf_path.stem)


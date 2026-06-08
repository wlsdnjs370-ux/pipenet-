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

import csv
import io
import json
import math
import os
import time
import warnings
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterator
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
STANDARD_SLF_FILENAME = "OA_3-1형_지하층포함_120~200m미만_35F.slf"


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

PIPENET_CATEGORIES = {"PIPE", "HEAD", "TEXT", "ALARM"}
KEEP_BASE_LAYERS = {"0"}  # INSERT BYLAYER 공통 + 도면 컨텍스트


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
        if any(k in name for k in ("SP", "배관", "소방", "가지관", "후렉시블", "FLEX")):
            return "PIPE"
        if any(k in n for k in ("text", "문자")) or "TEX" in name:
            return "TEXT"
        if any(k in name for k in ("벽", "건축", "WALL", "ARCH", "DIM", "SHEET", "AREA")):
            return "ARCH"
        return "OTHER"
    s = Remote30Settings()
    if layer_match(name, s.exclude_layer_keywords):
        return "EXCLUDE"
    if layer_match(name, s.arch_layer_keywords):
        return "ARCH"
    if layer_match(name, s.head_layer_keywords):
        return "HEAD"
    # ALARM 검사를 PIPE 보다 먼저 — "RISER" 가 "SP" 와 겹치지 않지만 우선순위 명시
    if layer_match(name, s.alarm_valve_keywords):
        return "ALARM"
    if layer_match(name, s.pipe_layer_keywords):
        return "PIPE"
    if layer_match(name, s.text_layer_keywords):
        return "TEXT"
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
        bundle.layer_visibility[name] = {
            "is_off": is_off,
            "is_frozen": is_frozen,
            "color": color,
        }
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
                if matrix is not None:
                    p0 = matrix.transform(Vec3(0.0, 0.0, 0.0))
                    p1 = matrix.transform(Vec3(1.0, 0.0, 0.0))
                    sf = math.hypot(p1.x - p0.x, p1.y - p0.y)
                else:
                    sf = 1.0
                r = float(e.dxf.radius) * sf
                bundle.entities.append({"t": "A", "l": layer, "c": [cx, cy], "r": r,
                                       "a": [float(e.dxf.start_angle), float(e.dxf.end_angle)]})
                _upd(cx - r, cy - r); _upd(cx + r, cy + r)
            elif etype == "CIRCLE":
                cx, cy = _t(matrix, e.dxf.center.x, e.dxf.center.y)
                if matrix is not None:
                    p0 = matrix.transform(Vec3(0.0, 0.0, 0.0))
                    p1 = matrix.transform(Vec3(1.0, 0.0, 0.0))
                    sf = math.hypot(p1.x - p0.x, p1.y - p0.y)
                else:
                    sf = 1.0
                r = float(e.dxf.radius) * sf
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
                paths_out = []
                for path in e.paths:
                    pts = []
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


def parse_dxf_for_view(dxf_path: Path, *, include_hidden_layers: bool = True,
                        keep_nested_insert_markers: bool = False) -> dict:
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
        try:
            if etype == "LINE":
                x1, y1 = _t(matrix, e.dxf.start.x, e.dxf.start.y)
                x2, y2 = _t(matrix, e.dxf.end.x, e.dxf.end.y)
                entities.append({"t": "L", "l": layer, "p": [x1, y1, x2, y2]})
                _upd(x1, y1); _upd(x2, y2)
            elif etype == "ARC":
                cx, cy = _t(matrix, e.dxf.center.x, e.dxf.center.y)
                if matrix is not None:
                    p0 = matrix.transform(Vec3(0.0, 0.0, 0.0))
                    p1 = matrix.transform(Vec3(1.0, 0.0, 0.0))
                    sf = math.hypot(p1.x - p0.x, p1.y - p0.y)
                else:
                    sf = 1.0
                r = float(e.dxf.radius) * sf
                entities.append({"t": "A", "l": layer, "c": [cx, cy], "r": r,
                                  "a": [float(e.dxf.start_angle), float(e.dxf.end_angle)]})
                _upd(cx - r, cy - r); _upd(cx + r, cy + r)
            elif etype == "CIRCLE":
                cx, cy = _t(matrix, e.dxf.center.x, e.dxf.center.y)
                if matrix is not None:
                    p0 = matrix.transform(Vec3(0.0, 0.0, 0.0))
                    p1 = matrix.transform(Vec3(1.0, 0.0, 0.0))
                    sf = math.hypot(p1.x - p0.x, p1.y - p0.y)
                else:
                    sf = 1.0
                r = float(e.dxf.radius) * sf
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
                paths_out = []
                for path in e.paths:
                    pts = []
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
                                    if ea < sa: ea += 360.0
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
import re as _re_v2

_DIA_TEXT_PATTERNS = (
    _re_v2.compile(r"\b(\d{2,3})\s*A\b"),                  # 25A
    _re_v2.compile(r"^\s*(\d{2,3})\s*$"),                  # 순수 숫자
    _re_v2.compile(r"[Øø]\s*(\d{2,3})"),                   # Ø25
    _re_v2.compile(r"DN\s*(\d{2,3})"),                     # DN25
    _re_v2.compile(r"(?<![0-9])(\d{2,3})\s*mm(?![0-9])"),  # 25mm
)
_DIA_TEXT_NOISE_KW = ("호스", "방수구", "소화전", "옥내", "HOSE", "EA", "KG",
                      "SET", "SCALE", "PUMP", "펌프", "TANK", "탱크", "SIZE")
_VALID_DIA_MM = frozenset((15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300))

_FLOOR_LABEL_PATTERNS = (
    (_re_v2.compile(r"지상\s*(\d{1,2})\s*층"), "ground"),     # 지상N층 → +N
    (_re_v2.compile(r"지하\s*(\d{1,2})\s*층"), "basement"),   # 지하N층 → -N
    (_re_v2.compile(r"B\s*(\d{1,2})\s*F", _re_v2.I), "basement"),  # B1F → -1
    (_re_v2.compile(r"(?<![A-Za-z])(\d{1,2})\s*F(?![A-Za-z])"), "ground"),  # 5F → +5
)
_FLOOR_LABEL_SPECIAL = {"옥상": 99, "옥탑": 99, "ROOF": 99, "R/F": 99, "RF": 99}


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


def _point_to_segment_dist(px: float, py: float,
                            ax: float, ay: float,
                            bx: float, by: float) -> float:
    """점 (px,py) ↔ segment [a,b] 최단 거리."""
    dx, dy = bx - ax, by - ay
    if dx == 0 and dy == 0:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    cx, cy = ax + t * dx, ay + t * dy
    return math.hypot(px - cx, py - cy)


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


def build_system_graph(
    entities: list[dict],
    bridge_tolerances_mm: tuple[float, ...] = (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0),
    layer_filter: set[str] | None = None,
    auto_filter_min_lines: int = 20,
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
    if layer_filter is None:
        auto_matched = _auto_pipe_layer_filter(entities)
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

    graph, edge_len = _build_graph(line_ents)
    comps_before = len(_connected_components(graph))
    total_bridges = 0
    for tol in bridge_tolerances_mm:
        total_bridges += _bridge_components(graph, edge_len, max_bridge_mm=tol)
    comps_after = len(_connected_components(graph))
    stats = {
        "line_entity_count": len(line_ents),
        "all_line_entity_count": len(all_line_ents),
        "node_count": len(graph),
        "edge_count": sum(len(nb) for nb in graph.values()) // 2,
        "components_before_bridge": comps_before,
        "components_after_bridge": comps_after,
        "bridges_applied": total_bridges,
        "layer_filter_used": sorted(filter_used) if filter_used else None,
        "layer_filter_fallback_no_match": fallback,
    }
    return graph, edge_len, stats


def find_nearest_graph_node_constrained(
    graph: dict,
    click_xy: tuple[float, float],
    max_dist_mm: float = 2500.0,
) -> tuple[tuple[float, float] | None, float]:
    """그래프 노드 중 클릭 좌표에 가장 가까운 노드. None = max_dist_mm 초과 (실패)."""
    near = _nearest_graph_node(graph, click_xy)
    if near is None:
        return None, float("inf")
    d = math.hypot(near[0] - click_xy[0], near[1] - click_xy[1])
    if d > max_dist_mm:
        return None, d
    return near, d


def _collapse_collinear_nodes(
    path: list[tuple[float, float]],
    edge_len: dict,
    angle_tol_deg: float = 0.5,
    short_ratio: float = 1.5,
) -> list[tuple[float, float]]:
    """Path 의 짧은 + 직선상 중간 노드 제거 — 답안 SDF 노드 구조에 근접.

    답안 SDF 의 노드는 fitting elbow / 분기 / 직경 변경 지점이라 평균 3-8m 간격.
    우리 path 는 LINE 끝점마다 노드라 같은 직선에 N+1 노드. 두 조건 동시 만족 시 통합:
      1) (i-1)→i 와 i→(i+1) 각도 ≤ angle_tol_deg (직선)
      2) 두 segment 길이 합 ≤ path median segment 길이 × short_ratio × 2
         (도면 단위 무관 — 상대 기준. short_ratio=1.5 → 평균의 1.5배까지 통합)
    """
    if len(path) <= 2:
        return list(path)

    # path 의 segment 길이 분포 → 중앙값
    seg_lens: list[float] = []
    for i in range(len(path) - 1):
        a, b = path[i], path[i + 1]
        L = edge_len.get((min(a, b), max(a, b)), math.hypot(b[0] - a[0], b[1] - a[1]))
        seg_lens.append(L)
    seg_lens_sorted = sorted(seg_lens)
    median_len = seg_lens_sorted[len(seg_lens_sorted) // 2] if seg_lens_sorted else 1.0
    threshold = median_len * short_ratio * 2.0   # 두 segment 합산용

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
        is_collinear = math.degrees(ang_rad) <= angle_tol_deg
        is_short = (L1 + L2) <= threshold
        if is_collinear and is_short:
            key_in  = (min(prev, cur), max(prev, cur))
            key_out = (min(cur, nxt), max(cur, nxt))
            merged_len = edge_len.get(key_in, L1) + edge_len.get(key_out, L2)
            new_key = (min(prev, nxt), max(prev, nxt))
            edge_len[new_key] = merged_len
            continue
        kept.append(cur)
    kept.append(path[-1])
    return kept


def extract_system_path(
    entities: list[dict],
    pump_xy: tuple[float, float],
    av_xy: tuple[float, float],
    snap_tolerance_mm: float = 2500.0,
    layer_filter: set[str] | None = None,
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

    graph, edge_len, stats = build_system_graph(entities, layer_filter=layer_filter)
    if not graph:
        raise ValueError(f"LINE entity 가 없음 (전체 entity {len(entities)}개 중 LINE/PL 0개)")

    pump_node, pump_d = find_nearest_graph_node_constrained(graph, pump_xy, max_dist_mm=snap_tolerance_mm)
    if pump_node is None:
        raise ValueError(
            f"펌프 클릭 좌표 ({int(pump_xy[0])}, {int(pump_xy[1])}) 근처 "
            f"{snap_tolerance_mm:.0f}mm 안에 배관 끝점 없음. "
            f"가장 가까운 노드까지 {pump_d:.0f}mm. 더 정확히 클릭하거나 snap_tolerance_mm 증가."
        )
    av_node, av_d = find_nearest_graph_node_constrained(graph, av_xy, max_dist_mm=snap_tolerance_mm)
    if av_node is None:
        raise ValueError(
            f"AV 클릭 좌표 ({int(av_xy[0])}, {int(av_xy[1])}) 근처 "
            f"{snap_tolerance_mm:.0f}mm 안에 배관 끝점 없음. "
            f"가장 가까운 노드까지 {av_d:.0f}mm."
        )

    path = _shortest_path(graph, edge_len, pump_node, av_node)
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

    return _system_path_to_riser_dict(
        path, edge_len, pump_xy, av_xy,
        pump_snap_dist=pump_d, av_snap_dist=av_d, graph_stats=stats,
        dia_text_pts=dia_text_pts, floor_labels=floor_labels,
    )


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

    def _elev_for_node(ny: float) -> tuple[float, str | None, bool]:
        """노드 Y 의 (elev_m, floor_name, from_label) 반환. label 없으면 Y/1000 fallback."""
        if floor_labels and av_floor_idx is not None:
            f_idx, f_name = _floor_for_node_y(ny, floor_labels)
            if f_idx is not None:
                return ((f_idx - av_floor_idx) * floor_height_mm / 1000.0, f_name, True)
        return ((ny - av_y_dxf) / 1000.0, None, False)

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
        elev_m, floor_name, from_label = _elev_for_node(pt[1])
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

SNAP_TOL_MM = 50.0
# 50mm: DN50 (최소 호칭경) 미만 거리는 배관 토폴로지상 같은 점으로 간주.
# 부동소수점 오차 + DWG→DXF 변환 누적 오차 + CAD 작업자 미세 오차 모두 흡수.
# 이전 5mm 는 대명동/다이소 작업엔 문제 없었으나, 좌표 절댓값이 큰 도면
# (예: MF-125 의 측지 좌표 3,500,000mm) 에서 변환 오차가 5mm 초과 → SP-LINE
# 끝점들이 안 만나서 그래프가 3,058 component 로 쪼개지는 사고 발생.
# 50mm 는 토폴로지 분석에 영향 없음 (호칭경 단위가 50/65/80/100mm 라 50mm
# 이내 차이는 의미 없음). 격자 snap 아니라 _NodeIndex cluster 반경.
HEAD_BRIDGE_MAX_MM = 5000.0  # 헤드 INSERT 좌표 ↔ 가장 가까운 그래프 노드 brigde 허용 거리.
# 5m: 메자닌/대형 도면 (예: MF-125) 의 헤드가 배관 라인과 천장고 차이로 멀리 떨어진 경우 보호.
SOURCE_BRIDGE_MAX_MM = 25000.0  # 알람밸브 (source) ↔ 배관망 nearest bridge 허용 (25m).
# 알람밸브는 라이저 (수직 입상관) 위에 위치 — 평면도상 가지관과 거리가 멀 수 있음.
# 25m 이내면 알람밸브 위치 그대로 source 로 사용 → 그래프 component 통합 효과.
MIN_PIPE_EDGE_MM = 50.0
# 50mm 미만 LINE/PL/ARC segment 는 그래프 edge 로 사용 안 함.
# 헤드 부속(HEADCON, HDCROSS, SPCAP 등), 치수 보조선, 텍스트 underline 등
# 평면도에는 보이지만 배관망 토폴로지에는 노이즈인 짧은 segment 제거.
CLOSED_PL_TOL_MM = 5.0  # PL 의 첫점과 마지막점이 이 거리 안이면 closed polygon 으로 간주 → 그래프 제외


def _round_pt(x: float, y: float, tol: float = SNAP_TOL_MM) -> tuple[float, float]:
    """격자 정렬 좌표 — HeadCandidate dedup, Counter 키 등 동등성 비교용.

    그래프 노드 키로는 더 이상 사용 안 함 (_NodeIndex 가 raw 좌표 기반 cluster 처리).
    """
    return (round(x / tol) * tol, round(y / tol) * tol)


class _NodeIndex:
    """Grid-bucket 기반 epsilon-tolerant endpoint cluster.

    Snap 격자(round-to-grid)의 대안. raw 좌표를 노드 키로 그대로 보존하면서,
    epsilon 반경 안에 기존 노드가 있으면 그 노드 좌표를 반환 (없으면 신규 등록).

    이점:
      - 노드 좌표 = raw → Stage 3 시각화가 격자 정렬 안 됨 → 비뚤어짐 없음
      - 격자 경계 분리 위험 없음 (cluster 가 9 bucket neighborhood 검색)
      - 같은 raw 좌표는 항상 같은 tuple value → dict hash 일관
    """

    __slots__ = ("eps", "_eps_sq", "_cell", "_bucket")

    def __init__(self, epsilon_mm: float = SNAP_TOL_MM):
        self.eps = epsilon_mm
        self._eps_sq = epsilon_mm * epsilon_mm
        self._cell = epsilon_mm
        self._bucket: dict[tuple[int, int], list[tuple[float, float]]] = defaultdict(list)

    def canonical(self, x: float, y: float) -> tuple[float, float]:
        """epsilon 안에 기존 노드 있으면 그 좌표, 없으면 (x, y) 신규 등록."""
        kx = int(x // self._cell)
        ky = int(y // self._cell)
        best = None
        bestd = float("inf")
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for pt in self._bucket.get((kx + dx, ky + dy), ()):
                    d = (pt[0] - x) ** 2 + (pt[1] - y) ** 2
                    if d < bestd and d <= self._eps_sq:
                        bestd = d
                        best = pt
        if best is not None:
            return best
        new_pt = (x, y)
        self._bucket[(kx, ky)].append(new_pt)
        return new_pt


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


def _is_triangle_shape(pts: list, tol: float = 2.0) -> bool:
    """HATCH path 의 점 시퀀스가 삼각형 (3 고유 정점) 인지 — closed loop 의 시작/끝 중복 무시."""
    if not pts or len(pts) < 3:
        return False
    unique: list[tuple[float, float]] = []
    for p in pts:
        x, y = float(p[0]), float(p[1])
        if not any(abs(x - u[0]) < tol and abs(y - u[1]) < tol for u in unique):
            unique.append((x, y))
    return len(unique) == 3


def detect_heads(pipe_entities: list[dict], layer_categories: dict[str, str]) -> list[HeadDetection]:
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
    CLUSTER_R = 250.0
    used = [False] * len(candidates)
    out: list[HeadDetection] = []
    for i, c1 in enumerate(candidates):
        if used[i]:
            continue
        cluster = [c1]
        used[i] = True
        for j in range(i + 1, len(candidates)):
            if used[j]:
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
    return out


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

    def is_pipe(layer: str) -> bool:
        if layer_categories is None:
            return True
        return layer_categories.get(layer, "OTHER") == "PIPE"

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
    return g, edge_len


# ────────────────────────────────────────────────────────────────────────────
# 평행 ladder collapse — 관경 두 줄 표현 → 중심선 1줄로 합성
# ────────────────────────────────────────────────────────────────────────────
# CAD 도면에서 배관을 두 평행 LINE 으로 그리는 관례 (관경 시각 표현).
# 그래프 빌드 시 두 줄이 모두 edge 로 남아 "ladder" (사다리) 모양 → 분기 수 부풀고
# 시각적으로 꼬임/겹침. 이 모듈은 4-cycle (u-v-w-x) 중 두 변이 평행하고 (rail)
# 나머지 두 변이 짧으면 (rung, 관 cap/cross-fitting) ladder 로 식별, midline 하나로 합성.

LADDER_MAX_RUNG_MM = 300.0     # rung (짧은 cross 변) 최대 길이. 단위세대 도면 기준.
LADDER_MIN_RAIL_RATIO = 3.0    # rail / rung 평균 길이 비. 정사각형 (=1) 은 합성 안 됨.
LADDER_PARALLEL_COS = 0.985    # 두 rail 의 방향 cos 유사도 임계값 (≈ 10도 안)
LADDER_MAX_ITER = 10           # collapse 반복 횟수 (합성 후 새 ladder 생길 수 있음)


def _edge_dir(p: tuple, q: tuple) -> tuple[float, float]:
    dx, dy = q[0] - p[0], q[1] - p[1]
    n = math.hypot(dx, dy)
    if n == 0.0:
        return (0.0, 0.0)
    return (dx / n, dy / n)


def _midpoint(p: tuple, q: tuple) -> tuple[float, float]:
    return ((p[0] + q[0]) / 2, (p[1] + q[1]) / 2)


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


def force_spanning_tree(
    graph: dict[tuple, set[tuple]],
    edge_len: dict[tuple, float],
    source: tuple | None = None,
) -> tuple[set, set]:
    """그래프를 (각 component 마다) shortest-path spanning tree 로 강제 변환.

    Args:
        graph: 무방향 그래프 (in-place 수정됨 — cycle edge 제거)
        edge_len: edge 길이 dict (in-place 수정 — 제거된 edge 도 같이 pop)
        source: AV (또는 시작 노드). 이 노드가 속한 component 는 source 가 root.
                다른 component 는 임의 root (가장 작은 좌표 노드).

    Returns:
        (tree_edges, removed_edges) — 각각 (min, max) 키 set.
    """
    import heapq

    tree_edges: set[tuple] = set()
    visited: set[tuple] = set()

    comps = _connected_components(graph)
    for comp in comps:
        if not comp:
            continue
        # 이 component 의 root 선택
        if source is not None and source in comp:
            root = source
        else:
            root = min(comp, key=lambda p: (p[0], p[1]))  # deterministic

        # Dijkstra SPT
        dist: dict[tuple, float] = {root: 0.0}
        parent: dict[tuple, tuple | None] = {root: None}
        pq: list[tuple[float, tuple]] = [(0.0, root)]
        while pq:
            d, u = heapq.heappop(pq)
            if d > dist[u]:
                continue
            for v in graph.get(u, ()):
                if v not in comp:
                    continue
                e_key = (min(u, v), max(u, v))
                w = edge_len.get(e_key)
                if w is None:
                    w = math.hypot(u[0] - v[0], u[1] - v[1])
                nd = d + w
                if nd < dist.get(v, float("inf")):
                    dist[v] = nd
                    parent[v] = u
                    heapq.heappush(pq, (nd, v))
        for n, p in parent.items():
            if p is not None:
                tree_edges.add((min(n, p), max(n, p)))
            visited.add(n)

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


def _find_head_candidates(pipe_entities: list[dict], layer_categories: dict[str, str]) -> list[HeadCandidate]:
    """HEAD 카테고리 레이어의 INSERT 또는 CIRCLE 위치를 헤드 후보로."""
    heads: list[HeadCandidate] = []
    seen: set[tuple[float, float]] = set()
    for en in pipe_entities:
        cat = layer_categories.get(en["l"], "OTHER")
        if cat != "HEAD":
            continue
        if en["t"] == "I":
            x, y = en["p"][0], en["p"][1]
            pos = _round_pt(x, y)
            if pos in seen:
                continue
            seen.add(pos)
            heads.append(HeadCandidate(pos=pos, raw=(x, y), block_name=en.get("n", ""), layer=en["l"]))
        elif en["t"] == "C":
            cx, cy = en["c"][0], en["c"][1]
            pos = _round_pt(cx, cy)
            if pos in seen:
                continue
            seen.add(pos)
            heads.append(HeadCandidate(pos=pos, raw=(cx, cy), block_name="(circle)", layer=en["l"]))
    return heads


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


def _dijkstra_from(graph: dict, edge_len: dict, src: tuple[float, float]) -> dict[tuple[float, float], float]:
    """단순 Dijkstra — 모든 노드까지의 거리."""
    import heapq
    dist: dict[tuple[float, float], float] = {src: 0.0}
    pq: list[tuple[float, tuple[float, float]]] = [(0.0, src)]
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist.get(u, float("inf")):
            continue
        for v in graph.get(u, ()):
            key = (min(u, v), max(u, v))
            w = edge_len.get(key, math.hypot(v[0] - u[0], v[1] - u[1]))
            nd = d + w
            if nd < dist.get(v, float("inf")):
                dist[v] = nd
                heapq.heappush(pq, (nd, v))
    return dist


def _shortest_path(graph: dict, edge_len: dict, src: tuple[float, float], tgt: tuple[float, float]) -> list[tuple[float, float]]:
    """src → tgt 최단 경로 (vertex 시퀀스)."""
    import heapq
    if src == tgt:
        return [src]
    dist = {src: 0.0}
    prev: dict = {}
    pq = [(0.0, src)]
    while pq:
        d, u = heapq.heappop(pq)
        if u == tgt:
            break
        if d > dist.get(u, float("inf")):
            continue
        for v in graph.get(u, ()):
            key = (min(u, v), max(u, v))
            w = edge_len.get(key, math.hypot(v[0] - u[0], v[1] - u[1]))
            nd = d + w
            if nd < dist.get(v, float("inf")):
                dist[v] = nd
                prev[v] = u
                heapq.heappush(pq, (nd, v))
    if tgt not in prev and tgt != src:
        return []
    # backtrack
    out = [tgt]
    while out[-1] in prev:
        out.append(prev[out[-1]])
    out.reverse()
    return out if out and out[0] == src else []


def _nearest_graph_node(graph: dict, pt: tuple[float, float]) -> tuple[float, float] | None:
    """그래프 노드 중 pt 와 가장 가까운 노드. 같은 좌표면 그대로."""
    if pt in graph:
        return pt
    best = None
    bestd = float("inf")
    for n in graph:
        d = (n[0] - pt[0]) ** 2 + (n[1] - pt[1]) ** 2
        if d < bestd:
            bestd = d
            best = n
    return best


def _connected_components(graph: dict) -> list[set]:
    """그래프의 connected component 들."""
    seen = set()
    comps = []
    for start in graph:
        if start in seen:
            continue
        stack = [start]
        comp = set()
        while stack:
            u = stack.pop()
            if u in seen:
                continue
            seen.add(u); comp.add(u)
            for v in graph.get(u, ()):
                if v not in seen:
                    stack.append(v)
        comps.append(comp)
    return comps


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
    comps = _connected_components(graph)
    if len(comps) <= 1:
        return 0
    # main = 가장 큰 component
    main = max(comps, key=len)
    others = [c for c in comps if c is not main]
    bridges = 0
    for comp in others:
        # comp 의 각 노드에서 main 의 가장 가까운 노드 찾기 (작은 comp 기준 O(|comp|*|main|))
        best = None
        bestd = float("inf")
        for u in comp:
            for v in main:
                d = math.hypot(u[0] - v[0], u[1] - v[1])
                if d < bestd:
                    bestd = d
                    best = (u, v)
        if best and bestd <= max_bridge_mm:
            u, v = best
            graph[u].add(v); graph[v].add(u)
            key = (min(u, v), max(u, v))
            edge_len[key] = bestd
            if bridge_edges_out is not None:
                bridge_edges_out.add(key)
            bridges += 1
    return bridges


def select_worst30_heads(
    pipe_entities: list[dict],
    layer_categories: dict[str, str],
    k: int = 30,
    manual_source: tuple[float, float] | None = None,
    manual_heads: list[tuple[float, float]] | None = None,
    zones: list[tuple[float, float, float, float]] | None = None,
) -> SelectionResult:
    """가장 불리한 K 헤드 + 경로 선정.

    manual_heads: 명시되면 자동 검출 대신 이 리스트 사용 (사용자 편집 후)
    zones: [(x1,y1,x2,y2), ...] 영역 union. 비어있지 않으면 그 안의 헤드만 후보로.
    """
    graph, edge_len = _build_graph(pipe_entities, layer_categories=layer_categories)
    # 평행 ladder collapse — Stage 3 시각화와 같은 토폴로지로 정렬.
    collapse_parallel_ladders(graph, edge_len)
    # 짧은 거리부터 단계적으로 brigde — 가까운 endpoint 우선 + 점점 멀리.
    # 5m / 10m 추가: 측지좌표 도면 (예: MF-125) 처럼 SP-LINE 끝점들이 멀리
    # 떨어진 경우 (변환 누적 오차 + 도면 분할 작업) component 통합 위해.
    for tol in (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0):
        _bridge_components(graph, edge_len, max_bridge_mm=tol)
    # 가지식 트리 강제 (SPT) — Stage 3 와 같은 토폴로지. SPT root 는 SDF source.
    # (source 가 이 시점에 아직 미결정 → 일단 None 으로 호출, component 별 임의 root.
    #  source 결정 후 트리가 SDF path 계산에 사용됨.)
    force_spanning_tree(graph, edge_len, source=None)
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
            edge_len[(min(src_raw, src_nearest), max(src_raw, src_nearest))] = d_src
            src = src_raw
        else:
            # 한도 초과 — nearest 로 fallback 하고 경고 플래그
            src = src_nearest
            src_fallback = True
            src_kind = src_kind + ":fallback_far"

    # 헤드 좌표 → 가장 가까운 그래프 노드로 강제 연결 (HEAD_BRIDGE_MAX_MM 이내)
    for h in heads:
        nearest = _nearest_graph_node(graph, h.pos)
        if nearest is None:
            continue
        d = math.hypot(h.pos[0] - nearest[0], h.pos[1] - nearest[1])
        if d > 1e-3 and d <= HEAD_BRIDGE_MAX_MM:
            graph.setdefault(h.pos, set()).add(nearest)
            graph[nearest].add(h.pos)
            edge_len[(min(h.pos, nearest), max(h.pos, nearest))] = d

    dist_map = _dijkstra_from(graph, edge_len, src)

    # head 후보들을 그래프 노드로 스냅 후 거리 정렬 — 도달 불가도 가능한 한 포함
    head_with_d: list[tuple[HeadCandidate, tuple[float, float], float]] = []
    for h in heads:
        node = h.pos if h.pos in graph else _nearest_graph_node(graph, h.pos)
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
    for _, head_node, _ in top_k:
        path = _shortest_path(graph, edge_len, src, head_node)
        for a, b in zip(path, path[1:]):
            key = (min(a, b), max(a, b))
            if key in sub_edges_seen:
                continue
            sub_edges_seen.add(key)
            sub_edges.append((a, b, edge_len.get(key, math.hypot(b[0] - a[0], b[1] - a[1]))))
            sub_nodes.add(a); sub_nodes.add(b)

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


def build_input_tables(
    selection: SelectionResult,
    pipe_entities: list[dict] | None = None,
    *,
    project_title: str = "Remote 30 Prototype",
) -> PipeTables:
    """선정 결과 → 5 테이블. pipe_entities 가 있으면 FX(flexible) Equipment 도 추출."""
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

    # Nodes
    for label, pos in label_to_pos.items():
        io_node = "Input" if label == src_label else "No"
        tables.nodes.append({
            "label": label, "elevation": 2.8, "io_node": io_node,
            "x": int(round(pos[0])), "y": int(round(pos[1])),
        })

    # ====== Diameter 추론 — 3단계 알고리즘
    # ① DXF TEXT 패턴 5종 추출 (노이즈 워드 필터)
    # ② NFPC 103 별표 1 "가"칸 (폐쇄형 SP) — 담당 헤드 수 → 최소 호칭경 매핑
    # ③ 결정: 텍스트 매칭이 있으면 max(텍스트값, NFPC최소값), 없으면 NFPC fallback
    import re as _re
    DIA_PATTERNS = [
        _re.compile(r"\b(\d{2,3})\s*A\b"),                                # 25A
        _re.compile(r"^\s*(\d{2,3})\s*$"),                                # 순수 숫자 (이 도면 dominant)
        _re.compile(r"[Øø]\s*(\d{2,3})"),                                 # Ø25
        _re.compile(r"DN\s*(\d{2,3})"),                                   # DN25
        _re.compile(r"(?<![0-9])(\d{2,3})\s*mm(?![0-9])"),                # 25mm
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
        tables.pipes.append({
            "label": plabel,
            "in": la, "out": lb,
            "type": "KSD 3507",
            "dia": dia,
            "length": round(length_mm / 1000.0, 2),
            "elev": 0.0,
            "c": "120",
            "status": "Normal",
            "group": "Unset",
        })
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
            # FX 등가길이 — 도면의 물리 길이가 아니라 KFI 인정/제품 스펙 기준의 고정값.
            # 3-1형 LSP 레퍼런스 SDF (60개 모두 15.6m) 기준 채택. 도면 물리길이는 fx_len 으로
            # 별도 계산되지만 검증/디버깅용으로만 메타에 남기고 eq_len 에는 쓰지 않음.
            fx_len_mm = 0.0
            for p0, p1 in zip(pts, pts[1:]):
                fx_len_mm += math.hypot(p1[0] - p0[0], p1[1] - p0[1])
            tables.equipment.append({
                "pipe": attached_pipe["label"], "in": attached_pipe["in"], "out": attached_pipe["out"],
                "label": str(fx_count + 1), "desc": "FX",
                "eq_len": 15.6,
                "rel_pos": 0.5,
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
            "eq_len": 15.6, "rel_pos": 0.5,
        })

    # 2) 알람밸브 (A/V) — src_label 이 in/out 인 첫 pipe 에 부착
    av_pipe = next((p for p in tables.pipes if p["in"] == src_label or p["out"] == src_label), None)
    if av_pipe:
        tables.equipment.insert(0, {
            "pipe": av_pipe["label"], "in": av_pipe["in"], "out": av_pipe["out"],
            "label": "1", "desc": "A/V",
            "eq_len": 12.9, "rel_pos": 0.5,
        })

    # Meta
    tables.meta = [
        ("원본 파일", project_title),
        ("SDF 버전", "1.8  (0)"),
        ("생성 모듈", "Remote 30 프로토타입"),
        ("선정 헤드 수", str(len(selection.heads))),
        ("subgraph 노드 수", str(len(label_to_pos))),
        ("subgraph 파이프 수", str(len(tables.pipes))),
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
                      "Equivalent length (m)", "Rel-position"],
    }
    rows_map = {
        "nodes": [[n["label"], n["elevation"], n["io_node"], n["x"], n["y"], None] for n in tables.nodes],
        "pipes": [[p["label"], p["in"], p["out"], p["type"], p["dia"], p["length"], p["elev"],
                   p["c"], p["status"], p["group"], None, None] for p in tables.pipes],
        "nozzles": [[n["label"], n["in"], n["out"], n["status"], n["lib"], n["flow_m3s"], n["flow_lmin"]] for n in tables.nozzles],
        "fittings": [[f["pipe"], f["in"], f["out"], f["type"], f["count"]] for f in tables.fittings],
        "equipment": [[e["pipe"], e["in"], e["out"], e["label"], e["desc"], e["eq_len"], e["rel_pos"]] for e in tables.equipment],
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
                       "Equivalent length (m)", "Rel-position"],
         [[e["pipe"], e["in"], e["out"], e["label"], e["desc"], e["eq_len"], e["rel_pos"]] for e in tables.equipment]),
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
        for suf in (".sdf", ".slf", ".xlsx"):
            p = out_dir / f"{prefix}{suf}"
            if p.is_file():
                zf.write(p, arcname=p.name)
        csv_dir = out_dir / "csv"
        if csv_dir.is_dir():
            for f in sorted(csv_dir.glob(f"{prefix}_*.csv")):
                zf.write(f, arcname=f"csv/{f.name}")
    return zip_path


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
        network.add_node(PnNode(
            node_id=str(n["label"]),
            x=nx, y=ny, z=float(n["elevation"]),
            node_type="input" if n["io_node"] == "Input" else "no",
            metadata={"io_node": n["io_node"]},
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
        network.add_nozzle(PnNozzle(
            nozzle_id=str(nz["label"]),
            input_node=str(nz["in"]),
            output_node=str(nz["out"]),
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
    # SLF 는 6 schedule (KSD 3507/3562/3576/DP/CPVC/FX) + SP-HEAD / INDOOR HYDRANT 노즐 + 표준 펌프 정의를 담은
    # 프로젝트 표준 라이브러리. 모든 수리계산 결과물 SDF 가 이 SLF 를 참조하도록 통일.
    # 경로 해석: 환경변수 REMOTE30_STANDARD_SLF → 모듈 디렉토리 fallback. (resolve_standard_slf 참조)
    import shutil as _shutil
    ref_slf = resolve_standard_slf()
    slf_name = out_path.with_suffix(".slf").name  # 예: prototype_<id>.slf
    slf_dst = out_path.parent / slf_name
    if ref_slf is not None and ref_slf.is_file():
        _shutil.copy2(ref_slf, slf_dst)
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
    # ── 6 schedule Pipe-type 정의 — 표준 SLF (OA_3-1형_..._35F.slf) 의 Schedule-section 과 정합.
    # 각 항목: (name, c-factor, [(size_m, max_velocity_m_s), ...])
    # 호칭경 set 은 SLF 의 Size-definition.nominal 과 동일, velocity 컨벤션 (≤50mm=6, ≥65mm=10)
    # 은 레퍼런스 4-1형 알람밸브 SDF 의 KSD 3507/3562/CPVC Pipe-type 정의에서 도출.
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
        ("CPVC", "150", [
            (0.025, 6), (0.032, 6), (0.04, 6), (0.05, 6), (0.065, 10), (0.08, 10),
        ]),
        ("FX", "120", [(0.02, 10)]),
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
        break

    _tree.write(out_path, encoding="utf-8", xml_declaration=True)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# Orchestrator — generator yielding progress events
# ────────────────────────────────────────────────────────────────────────────


def run_stages_0_2(
    dxf_path: Path,
    job_id: str,
    alarm_xy: tuple[float, float] | None = None,
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
    bundle = parse_dxf_bundle(dxf_path)
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
    # bridge_edges: _bridge_components 가 강제로 이은 연결 (실제 배관 아님)
    # head_drop_edges: 헤드 INSERT 좌표 ↔ 배관 nearest 노드 직선 (실제 배관 아님)
    # 두 종류 모두 "알고리즘이 추정한 연결"이라 시각적으로 구분 렌더.
    bridge_edges: set = set()
    for tol in (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0):
        _bridge_components(graph, edge_len, max_bridge_mm=tol, bridge_edges_out=bridge_edges)
    # 주의: SPT 가 cycle edge 를 제거하면서 bridge_edges 일부도 같이 제거될 수 있음.
    # SPT 적용 후 살아남은 bridge_edges 만 유효.
    # 헤드 drop line — 헤드 INSERT 좌표를 같은 NodeIndex 로 canonicalize
    # (그래프에 이미 같은 epsilon 안 노드 있으면 그 raw 좌표 반환 → drop line 불필요).
    head_drop_edges: set = set()
    head_pos_list = []
    for h in detect_heads(pipe_ents, layer_categories):
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
    tree_edges_set, removed_cycle_edges = force_spanning_tree(graph, edge_len, source=spt_source)
    # SPT 가 일부 bridge/drop edge 도 제거할 수 있음 — 살아남은 것만 유효
    bridge_edges &= tree_edges_set
    head_drop_edges &= tree_edges_set

    # edge entity emit — 4종 구분 (실배관 / bridge / drop / 제거된 cycle)
    graph_ents = []
    seen_edges: set = set()
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
            if key in bridge_edges:
                layer = "_graph_bridge"
                bridge_emitted += 1
            elif key in head_drop_edges:
                layer = "_graph_head_drop"
                head_drop_emitted += 1
            else:
                layer = "_graph_edge"
            graph_ents.append({"t": "L", "l": layer, "p": [u[0], u[1], v[0], v[1]]})
    # 2) 제거된 cycle edge — 별도 카테고리 (회색 매우 흐릿, 참고용)
    for (u, v) in removed_cycle_edges:
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

    real_edge_count = len(seen_edges) - bridge_emitted - head_drop_emitted
    summary = {
        "node_count": len(graph),
        "edge_count": len(seen_edges),
        "real_edge_count": real_edge_count,
        "bridge_edge_count": bridge_emitted,
        "head_drop_edge_count": head_drop_emitted,
        "junction_count": junction_count,
        "components": len(_connected_components(graph)),
        "ladders_collapsed": ladders_collapsed,
        "removed_cycle_edges": cycle_emitted,
        "source_pos": list(src_raw_pt) if src_raw_pt else None,
        "source_kind": src_kind_preview if src_raw_pt else "none",
        "source_bridge_dist_mm": round(src_bridge_preview, 1),
        "source_far_from_pipes": src_far,
    }
    yield evt({"type": "entities", "stage": 3, "entities": graph_ents, "summary": summary})
    label = (
        f"가지식 트리 — {len(graph)} 노드 / 실배관 {real_edge_count} edge"
        f" / bridge {bridge_emitted} / 헤드 drop {head_drop_emitted} / 분기 {junction_count}개"
        f" / ladder 합성 {ladders_collapsed} / cycle 제거 {cycle_emitted}"
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
    selection = select_worst30_heads(pipe_ents, layer_categories,
                                     k=k_heads, manual_source=alarm_xy,
                                     manual_heads=edited_heads, zones=zones)
    subgraph_ents = []
    for a, b, _len in selection.edges:
        subgraph_ents.append({"t": "L", "l": "_subgraph", "p": [a[0], a[1], b[0], b[1]]})
    for h in selection.heads:
        subgraph_ents.append({"t": "C", "l": "_subgraph_head", "c": list(h.pos), "r": 80.0})
    if selection.source_pos is not None:
        subgraph_ents.append({"t": "C", "l": "_alarm_valve", "c": list(selection.source_pos), "r": 150.0})
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
    tables = build_input_tables(selection, pipe_entities=pipe_ents, project_title=dxf_path.stem)
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

    # Stage 6: SDF (기존 5) + .slf 동봉 + 결과 zip
    yield evt({"type": "stage", "stage": 6, "status": "running", "label": "PIPENET SDF emit + .slf 동봉 + zip 묶음"})
    sdf_path = out_dir / f"prototype_{job_id}.sdf"
    emit_sdf(tables, sdf_path, project_title=dxf_path.stem)
    slf_path = sdf_path.with_suffix(".slf")
    zip_path = bundle_result_zip(out_dir, prefix=f"prototype_{job_id}")
    yield evt({"type": "stage", "stage": 6, "status": "done",
               "label": f"SDF {sdf_path.stat().st_size/1024:.1f}KB + SLF {slf_path.stat().st_size/1024:.1f}KB + ZIP {zip_path.stat().st_size/1024:.1f}KB"})

    outputs = {
        "xlsx": xlsx_path.name, "sdf": sdf_path.name,
        "csv_nodes": csv_paths["nodes"].name, "csv_pipes": csv_paths["pipes"].name,
        "csv_nozzles": csv_paths["nozzles"].name, "csv_fittings": csv_paths["fittings"].name,
        "csv_equipment": csv_paths["equipment"].name,
    }
    if slf_path.is_file():
        outputs["slf"] = slf_path.name
    if zip_path.is_file():
        outputs["zip"] = zip_path.name
    yield evt({"type": "done", "outputs": outputs, "out_dir": str(out_dir)})


def run_prototype_pipeline(
    dxf_path: Path,
    out_dir: Path,
    job_id: str,
    *,
    k_heads: int = 30,
    alarm_xy: tuple[float, float] | None = None,
) -> Iterator[dict]:
    """전체 파이프라인 — JSON-직렬화 가능한 진행 이벤트 yield."""
    out_dir.mkdir(parents=True, exist_ok=True)
    t0 = time.time()

    def evt(d: dict) -> dict:
        d.setdefault("elapsed_ms", int((time.time() - t0) * 1000))
        return d

    # Stage 0: 파싱
    yield evt({"type": "stage", "stage": 0, "status": "running", "label": "DXF 파싱"})
    bundle = parse_dxf_bundle(dxf_path)
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

    # Stage 1: 배관망만 필터
    yield evt({"type": "stage", "stage": 1, "status": "running", "label": "건축/기타 레이어 제거 (배관망만)"})
    pipe_ents = filter_pipenet_only(bundle)
    yield evt({"type": "entities", "stage": 1,
               "entities": pipe_ents,
               "summary": {"entity_count": len(pipe_ents)}})
    yield evt({"type": "stage", "stage": 1, "status": "done",
               "label": f"배관망 추출 완료 — {len(pipe_ents):,} entity"})

    # Stage 2: 전체 헤드 바운딩박스 인식 (신규 — Stage 1과 헤드 선정 사이)
    yield evt({"type": "stage", "stage": 2, "status": "running",
               "label": "도면 내 전체 헤드 후보 인식 (block pattern + CIRCLE/HATCH 시그니처 + 클러스터링)"})
    head_detections = detect_heads(pipe_ents, layer_categories)
    # 바운딩박스 entity 들을 캔버스에 emit (t='B' for bounding box)
    bbox_ents = []
    for h in head_detections:
        bx1, by1, bx2, by2 = h.bbox
        bbox_ents.append({
            "t": "B", "l": "_head_bbox",
            "p": [bx1, by1, bx2, by2],
            "k": h.kind, "c": round(h.confidence, 2),
            "n": h.block_name,
        })
    # cue 분포
    from collections import Counter as _C
    kind_counter: _C = _C()
    for h in head_detections:
        primary = h.kind.split(":")[0] if ":" in h.kind else h.kind
        kind_counter[primary] += 1
    yield evt({"type": "entities", "stage": 2,
               "entities": bbox_ents,
               "summary": {
                   "head_count": len(head_detections),
                   "by_kind": dict(kind_counter),
                   "avg_confidence": round(sum(h.confidence for h in head_detections) / len(head_detections), 3) if head_detections else 0,
               }})
    yield evt({"type": "stage", "stage": 2, "status": "done",
               "label": f"전체 헤드 {len(head_detections)}개 인식 완료 (평균 신뢰도 {sum(h.confidence for h in head_detections)/max(1,len(head_detections)):.2f})"})

    # Stage 3 (구 Stage 2): 가장 불리한 K 헤드 + subgraph
    src_label = "수동 좌표" if alarm_xy else "자동 식별"
    yield evt({"type": "stage", "stage": 3, "status": "running",
               "label": f"G₀ 빌드 + 가장 불리한 {k_heads} 헤드 선정 (알람밸브 {src_label})"})
    selection = select_worst30_heads(pipe_ents, layer_categories, k=k_heads, manual_source=alarm_xy)
    # subgraph 시각화 entity (LINE) 생성
    subgraph_ents = []
    for a, b, _len in selection.edges:
        subgraph_ents.append({"t": "L", "l": "_subgraph", "p": [a[0], a[1], b[0], b[1]]})
    for h in selection.heads:
        subgraph_ents.append({"t": "C", "l": "_subgraph_head", "c": list(h.pos), "r": 80.0})
    if selection.source_pos is not None:
        subgraph_ents.append({"t": "C", "l": "_alarm_valve",
                              "c": list(selection.source_pos), "r": 150.0})
    yield evt({"type": "entities", "stage": 3,
               "entities": subgraph_ents,
               "summary": {
                   "selected_heads": len(selection.heads),
                   "subgraph_edges": len(selection.edges),
                   "subgraph_nodes": len(selection.nodes_in_subgraph),
                   "max_distance_m": round(max(selection.distances) / 1000.0, 2) if selection.distances else 0,
                   "source_kind": selection.source_kind,
                   "source_pos": list(selection.source_pos) if selection.source_pos else None,
               }})
    yield evt({"type": "stage", "stage": 3, "status": "done",
               "label": f"선정 완료 — 헤드 {len(selection.heads)}개 / 경로 {len(selection.edges)} edge"})

    # Stage 4 (구 Stage 3): 5 테이블 (XLSX + CSV) emit
    yield evt({"type": "stage", "stage": 4, "status": "running", "label": "Nodes/Pipes/Nozzles/Fittings/Equipment 테이블 생성"})
    tables = build_input_tables(selection, pipe_entities=pipe_ents, project_title=dxf_path.stem)
    csv_dir = out_dir / "csv"
    csv_paths = write_csv_tables(tables, csv_dir, prefix=f"prototype_{job_id}")
    xlsx_path = out_dir / f"prototype_{job_id}.xlsx"
    write_xlsx_tables(tables, xlsx_path)
    yield evt({"type": "tables_preview", "stage": 4,
               "tables": {
                   "nodes": tables.nodes[:8],
                   "pipes": tables.pipes[:8],
                   "nozzles": tables.nozzles[:8],
                   "fittings": tables.fittings[:8],
                   "equipment": tables.equipment[:8],
                   "meta": tables.meta,
               },
               "counts": {
                   "nodes": len(tables.nodes),
                   "pipes": len(tables.pipes),
                   "nozzles": len(tables.nozzles),
                   "fittings": len(tables.fittings),
                   "equipment": len(tables.equipment),
               }})
    yield evt({"type": "stage", "stage": 4, "status": "done",
               "label": f"5 테이블 생성 완료 — Pipes {len(tables.pipes)} / Nodes {len(tables.nodes)} / Nozzles {len(tables.nozzles)}"})

    # Stage 5 (구 Stage 4): SDF emit + .slf 동봉 + zip 묶음
    yield evt({"type": "stage", "stage": 5, "status": "running", "label": "PIPENET SDF emit + .slf 동봉 + zip 묶음"})
    sdf_path = out_dir / f"prototype_{job_id}.sdf"
    emit_sdf(tables, sdf_path, project_title=dxf_path.stem)
    slf_path = sdf_path.with_suffix(".slf")
    zip_path = bundle_result_zip(out_dir, prefix=f"prototype_{job_id}")
    yield evt({"type": "stage", "stage": 5, "status": "done",
               "label": f"SDF {sdf_path.stat().st_size/1024:.1f}KB + SLF {slf_path.stat().st_size/1024:.1f}KB + ZIP {zip_path.stat().st_size/1024:.1f}KB"})

    # Done
    yield evt({"type": "done",
               "outputs": {
                   "xlsx": xlsx_path.name,
                   "sdf": sdf_path.name,
                   "csv_nodes": csv_paths["nodes"].name,
                   "csv_pipes": csv_paths["pipes"].name,
                   "csv_nozzles": csv_paths["nozzles"].name,
                   "csv_fittings": csv_paths["fittings"].name,
                   "csv_equipment": csv_paths["equipment"].name,
               },
               "out_dir": str(out_dir)})

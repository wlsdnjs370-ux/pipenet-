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

PIPENET_CATEGORIES = {"PIPE", "HEAD", "TEXT"}
KEEP_BASE_LAYERS = {"0"}  # INSERT BYLAYER 공통 + 도면 컨텍스트


def _categorize_layer(name: str) -> str:
    """Remote30Settings 기준 layer 카테고리. 가능하면 외부 모듈 사용."""
    if Remote30Settings is None or layer_match is None:
        # fallback heuristic
        n = name.lower()
        if any(k in n for k in ("소화기", "옥내소화전", "자동식", "co2")):
            return "EXCLUDE"
        if any(k in name for k in ("HEAD", "헤드", "SP-H", "하향식", "상향식", "헤드반경")):
            return "HEAD"
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
    if layer_match(name, s.pipe_layer_keywords):
        return "PIPE"
    if layer_match(name, s.text_layer_keywords):
        return "TEXT"
    return "OTHER"


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

    bbox = [float("inf"), float("inf"), float("-inf"), float("-inf")]

    def _upd(x: float, y: float) -> None:
        if x < bbox[0]:
            bbox[0] = x
        if y < bbox[1]:
            bbox[1] = y
        if x > bbox[2]:
            bbox[2] = x
        if y > bbox[3]:
            bbox[3] = y

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

    if bbox[0] == float("inf"):
        bbox = [0.0, 0.0, 1.0, 1.0]
    bundle.bbox = bbox

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

SNAP_TOL_MM = 200.0  # 200mm: 미세 segment 가 snap 단계에서 통합 (참조 ref 의 min pipe 길이가 320mm 이므로 200mm 까지 OK)
HEAD_BRIDGE_MAX_MM = 2000.0  # 헤드 INSERT 좌표 ↔ 가장 가까운 그래프 노드 brigde 허용 거리
SOURCE_BRIDGE_MAX_MM = 10000.0  # 알람밸브(source) ↔ 배관망 nearest 노드 bridge 허용 거리 (10m). 초과 시 nearest 로 fallback + 경고


def _round_pt(x: float, y: float, tol: float = SNAP_TOL_MM) -> tuple[float, float]:
    return (round(x / tol) * tol, round(y / tol) * tol)


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
    "A$C39172136", "A$C3F157AFD", "A$C60792707",
    "A$C6B5253FE", "A$C563427C5", "A$C324C7814", "A$C0F5C7CDB",
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


def _build_graph(pipe_entities: list[dict]) -> tuple[dict[tuple[float, float], set[tuple[float, float]]], dict[tuple, float]]:
    """파이프 LINE/PL 으로부터 무방향 그래프 빌드 (snap 적용)."""
    g: dict[tuple[float, float], set[tuple[float, float]]] = defaultdict(set)
    edge_len: dict[tuple, float] = {}
    for en in pipe_entities:
        # PIPE 카테고리만 그래프 구성에 사용 (LINE, PL)
        if en["t"] == "L":
            x1, y1, x2, y2 = en["p"]
            a = _round_pt(x1, y1); b = _round_pt(x2, y2)
            if a == b:
                continue
            g[a].add(b); g[b].add(a)
            key = (min(a, b), max(a, b))
            edge_len[key] = math.hypot(b[0] - a[0], b[1] - a[1])
        elif en["t"] == "PL":
            pts = en["p"]
            for p0, p1 in zip(pts, pts[1:]):
                a = _round_pt(p0[0], p0[1]); b = _round_pt(p1[0], p1[1])
                if a == b:
                    continue
                g[a].add(b); g[b].add(a)
                key = (min(a, b), max(a, b))
                edge_len[key] = math.hypot(b[0] - a[0], b[1] - a[1])
    return g, edge_len


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
    """알람밸브 자동 식별 — 4-tier fallback:
      1) block_name 에 'ALARM' 또는 '알람' 포함된 INSERT
      2) '배관-SP 2차' 또는 'SP 2차' 레이어의 첫 INSERT (입상→알람→가지 source)
      3) '배관-SP 2차' 레이어의 LINE 의 endpoint 중 가지관 그래프와 가장 가까운 점
      4) None (호출자가 fallback 처리)
    """
    for en in pipe_entities:
        if en["t"] != "I":
            continue
        bn = (en.get("n") or "").upper()
        if "ALARM" in bn or "알람" in bn:
            return _round_pt(en["p"][0], en["p"][1]), "alarm_block"
    for en in pipe_entities:
        if en["t"] != "I":
            continue
        if "배관-SP 2차" in en["l"] or "SP 2차" in en["l"]:
            return _round_pt(en["p"][0], en["p"][1]), "secondary_layer_insert"
    # 2차 배관 LINE 의 endpoint 들 수집
    secondary_endpoints: list[tuple[float, float]] = []
    for en in pipe_entities:
        if en["t"] == "L" and ("배관-SP 2차" in en["l"] or "SP 2차" in en["l"]):
            p = en["p"]
            secondary_endpoints.append(_round_pt(p[0], p[1]))
            secondary_endpoints.append(_round_pt(p[2], p[3]))
    if secondary_endpoints:
        # endpoint 중 가장 자주 등장하는 점 (T-junction with 가지관) — alarm valve 가 거기
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


def _bridge_components(graph: dict, edge_len: dict, max_bridge_mm: float = 500.0) -> int:
    """끊어진 component 들을 가장 가까운 endpoint 쌍 연결 — 50cm 이내만."""
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
            edge_len[(min(u, v), max(u, v))] = bestd
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
    graph, edge_len = _build_graph(pipe_entities)
    # 짧은 거리부터 단계적으로 brigde — 가까운 endpoint 우선 + 점점 멀리
    for tol in (200.0, 500.0, 1000.0, 2000.0):
        _bridge_components(graph, edge_len, max_bridge_mm=tol)
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
               "summary": {"entity_count": len(bundle.entities), "layer_count": len(bundle.layers)}})
    yield evt({"type": "stage", "stage": 0, "status": "done",
               "label": f"DXF 파싱 완료 — {len(bundle.entities):,} entity / {len(bundle.layers)} 레이어"})

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
    # snap + 컴포넌트 brigde + 헤드 drop line 모두 포함된 최종 그래프.
    yield evt({"type": "stage", "stage": 3, "status": "running",
               "label": "전체 배관망 그래프 인식 (snap + 컴포넌트 bridge + 헤드 drop line)"})
    graph, edge_len = _build_graph(pipe_ents)
    for tol in (200.0, 500.0, 1000.0, 2000.0):
        _bridge_components(graph, edge_len, max_bridge_mm=tol)
    # 헤드 drop line 추가 (select_worst30_heads 와 동일 로직)
    head_pos_list = []
    for h in detect_heads(pipe_ents, layer_categories):
        head_pos_list.append(_round_pt(h.pos[0], h.pos[1]))
    for hp in head_pos_list:
        nearest = _nearest_graph_node(graph, hp)
        if nearest is None or hp == nearest:
            continue
        d = math.hypot(hp[0] - nearest[0], hp[1] - nearest[1])
        if d > 1e-3 and d <= HEAD_BRIDGE_MAX_MM:
            graph.setdefault(hp, set()).add(nearest)
            graph[nearest].add(hp)
            edge_len[(min(hp, nearest), max(hp, nearest))] = d
    # edge entity 모두 같은 색 emit (사용자 옵션 C)
    graph_ents = []
    seen_edges: set = set()
    for u, neighbors in graph.items():
        for v in neighbors:
            key = (min(u, v), max(u, v))
            if key in seen_edges:
                continue
            seen_edges.add(key)
            graph_ents.append({"t": "L", "l": "_graph_edge", "p": [u[0], u[1], v[0], v[1]]})
    # junction 노드 (차수 ≥ 3) 만 점으로
    junction_count = 0
    for n, neighbors in graph.items():
        if len(set(neighbors)) >= 3:
            graph_ents.append({"t": "C", "l": "_graph_junction", "c": [n[0], n[1]], "r": 80.0})
            junction_count += 1

    # ── 알람밸브(source) 시각화 — 사용자 지정 좌표 또는 자동 식별
    if alarm_xy is not None:
        src_raw_pt = _round_pt(alarm_xy[0], alarm_xy[1])
        src_kind_preview = "manual"
    else:
        src_raw_pt, src_kind_preview = _find_source(pipe_ents, layer_categories)
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

    summary = {
        "node_count": len(graph),
        "edge_count": len(seen_edges),
        "junction_count": junction_count,
        "components": len(_connected_components(graph)),
        "source_pos": list(src_raw_pt) if src_raw_pt else None,
        "source_kind": src_kind_preview if src_raw_pt else "none",
        "source_bridge_dist_mm": round(src_bridge_preview, 1),
        "source_far_from_pipes": src_far,
    }
    yield evt({"type": "entities", "stage": 3, "entities": graph_ents, "summary": summary})
    label = f"배관망 그래프 — {len(graph)} 노드 / {len(seen_edges)} edge / 분기 {junction_count}개"
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
               "summary": {"entity_count": len(bundle.entities), "layer_count": len(bundle.layers)}})
    yield evt({"type": "stage", "stage": 0, "status": "done",
               "label": f"DXF 파싱 완료 — {len(bundle.entities):,} entity / {len(bundle.layers)} 레이어"})

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

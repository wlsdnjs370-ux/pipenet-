"""Remote 30 전체 배관망 총괄 모듈 (10번 모듈 백엔드).

기존 ``remote30_prototype`` 의 헤드망 추출 로직(Stage A)을 재사용하고,
zone별 라이저 템플릿(Stage B), 라이저↔헤드망 stitch(Stage C),
PIPENET-native 후처리 + Pump-fan/Elastomeric-valve 직렬화(Stage D) 를
추가하여 펌프 → 감압밸브 → 알람밸브 → 헤드 30개 전 구간 SDF 를 생성한다.

흐름::

    OverallInputs (DXF + ZoneSpec + (선택) BuildingPressureProfile)
            │
            ├── Stage A — run_stages_0_2 (remote30_prototype 재사용)
            │                 → PipeTables (헤드망)
            │
            ├── Stage B — build_riser(zone_spec, profile)
            │                 → RiserTables (펌프/PRV/라이저)
            │
            ├── Stage C — stitch_riser_and_heads(riser, head_tables)
            │                 → CombinedTables
            │
            └── Stage D — emit_full_sdf(combined, out_path)
                              → 완성 SDF + 동봉 SLF

신규 attribute (vs prototype 의 PipeTables)::

    RiserTables.pumps   — <Pump-fan> 직렬화용 dict 리스트
    RiserTables.valves  — <Elastomeric-valve> 직렬화용 dict 리스트
"""

from __future__ import annotations

import csv
import io
import json
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Iterator


# ────────────────────────────────────────────────────────────────────────────
# 상수
# ────────────────────────────────────────────────────────────────────────────

KGF_CM2_TO_PA = 98066.5    # 1 kg/cm² (kgf/cm²) → Pa
M_TO_PA = 9806.65          # 1 m 수두 → Pa (물 비중 1.0 기준)
ATM_PA = 101325.0          # 1 기압 (boundary condition)


# ────────────────────────────────────────────────────────────────────────────
# Zone 정의
# ────────────────────────────────────────────────────────────────────────────

class ZoneType(Enum):
    """처리 가능한 zone 타입 — 답안 SDF 의 라이저 구조에서 도출.

    압력분포표(예: 대명동 201동 PDF)의 각 행 → ZoneType 매핑::

        옥상 + 27F~49F          → HSP_PUMP        (펌프식 부스터)
        25F~26F                 → LSP_GRAVITY     (감압 없음, 자연낙차)
        2F~24F (24F 1차 PRV)    → LSP_1STAGE      (1차 감압)
        1F~B4 (1.5F 2차 PRV)    → LLSP_2STAGE     (1차 + 2차 감압)
    """
    HSP_PUMP = "hsp_pump"
    LSP_GRAVITY = "lsp_gravity"
    LSP_1STAGE = "lsp_1stage"
    LLSP_2STAGE = "llsp_2stage"


# ────────────────────────────────────────────────────────────────────────────
# 데이터 모델
# ────────────────────────────────────────────────────────────────────────────

@dataclass
class FloorRow:
    """압력분포표의 한 행 (옥상~B4 까지 1층 1행)."""
    floor_label: str          # "옥상층", "49층", "1층", "B1층" 등
    height_m: float           # 층고 (m)
    head_drop_m: float        # 누적 낙차 (m, 옥상 수원 기준)
    after_prv_m: float | None = None  # 감압 이후 수두 (m) — 감압 구간만
    note: str = ""            # 비고 ("자연낙차시작점", "1차 감압밸브 사용구간" 등)


@dataclass
class BuildingPressureProfile:
    """빌딩 전체 압력 흐름 표 — 옥상부터 최하층까지 1행씩.

    CSV/엑셀 업로드 또는 사용자 직접 입력 폼으로 생성. 없을 수도 있음 (선택적).
    """
    building_name: str = ""
    floors: list[FloorRow] = field(default_factory=list)

    def find_by_label(self, floor_label: str) -> FloorRow | None:
        for row in self.floors:
            if row.floor_label == floor_label:
                return row
        return None

    @classmethod
    def from_csv(cls, csv_path: Path, *, building_name: str = "") -> "BuildingPressureProfile":
        """CSV 파서 — 컬럼: floor_label, height_m, head_drop_m, after_prv_m, note.

        헤더 행은 한글/영문 모두 허용 (구분/층고/낙차압/감압이후/비고).
        """
        rows: list[FloorRow] = []
        # 한글 컬럼명 → 영문 키 매핑 (PDF 표 헤더에 맞춤)
        KCOL = {
            "구분": "floor_label", "층": "floor_label",
            "층고": "height_m",
            "낙차압": "head_drop_m", "낙차": "head_drop_m",
            "감압이후": "after_prv_m", "감압후": "after_prv_m",
            "비고": "note",
        }
        with csv_path.open("r", encoding="utf-8-sig", newline="") as fp:
            reader = csv.DictReader(fp)
            for raw in reader:
                # 한글 헤더면 매핑 적용
                norm: dict[str, Any] = {}
                for k, v in raw.items():
                    key = KCOL.get((k or "").strip(), (k or "").strip())
                    norm[key] = (v or "").strip()
                if not norm.get("floor_label"):
                    continue
                try:
                    rows.append(FloorRow(
                        floor_label=norm["floor_label"],
                        height_m=float(norm.get("height_m") or 0),
                        head_drop_m=float(norm.get("head_drop_m") or 0),
                        after_prv_m=(float(norm["after_prv_m"]) if norm.get("after_prv_m") else None),
                        note=norm.get("note", "") or "",
                    ))
                except (ValueError, KeyError):
                    continue
        return cls(building_name=building_name, floors=rows)

    @classmethod
    def from_xlsx(cls, xlsx_path: Path, *, sheet: str | int = 0,
                  building_name: str = "") -> "BuildingPressureProfile":
        """엑셀(.xlsx) 파서 — 첫 시트 또는 지정 시트의 같은 컬럼 구조."""
        from openpyxl import load_workbook
        wb = load_workbook(xlsx_path, read_only=True, data_only=True)
        ws = wb[sheet] if isinstance(sheet, str) else wb.worksheets[sheet]
        header_row = next(ws.iter_rows(values_only=True))
        KCOL = {
            "구분": "floor_label", "층": "floor_label",
            "층고": "height_m",
            "낙차압": "head_drop_m", "낙차": "head_drop_m",
            "감압이후": "after_prv_m", "감압후": "after_prv_m",
            "비고": "note",
        }
        col_idx = {KCOL.get((str(h) or "").strip(), (str(h) or "").strip()): i
                   for i, h in enumerate(header_row) if h is not None}
        rows: list[FloorRow] = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r:
                continue
            def _get(key, default=None):
                i = col_idx.get(key)
                return r[i] if i is not None and i < len(r) else default
            label = _get("floor_label")
            if not label:
                continue
            try:
                rows.append(FloorRow(
                    floor_label=str(label),
                    height_m=float(_get("height_m") or 0),
                    head_drop_m=float(_get("head_drop_m") or 0),
                    after_prv_m=(float(_get("after_prv_m")) if _get("after_prv_m") is not None else None),
                    note=str(_get("note") or ""),
                ))
            except (ValueError, TypeError):
                continue
        wb.close()
        return cls(building_name=building_name, floors=rows)


@dataclass
class ZoneSpec:
    """처리 대상 zone 의 사양. 사용자 입력 (라디오 + 폼).

    필드 의미::

        zone_type             — ZoneType (HSP_PUMP / LSP_GRAVITY / LSP_1STAGE / LLSP_2STAGE)
        target_floor          — 추출 대상 층 ("16층" — Stage A 헤드망과 동일)
        prv1_target_pa        — 1차 PRV 출구압 (Pa). 자연낙차 + 감압 zone 에서.
        prv2_target_pa        — 2차 PRV 출구압 (Pa). LLSP_2STAGE 만.
        pump_library_name     — Library-pump 의 SLF Pump-definition 이름.
                                기본 "SP_162M_2900LPM" (표준 SLF 와 정합).
        pump_count            — Pump-fan 개수 (HSP 보통 2개 = 1차+2차 부스터)

    압력분포표가 있으면 prv1/prv2 target 은 표에서 자동 도출 가능, 없으면 직접 입력.
    """
    zone_type: ZoneType
    target_floor: str = ""
    prv1_target_pa: float | None = None
    prv2_target_pa: float | None = None
    pump_library_name: str = "SP_162M_2900LPM"
    pump_count: int = 2


@dataclass
class OverallInputs:
    """모듈 10 의 입력 일체."""
    dxf_path: Path
    zone_spec: ZoneSpec
    profile: BuildingPressureProfile | None = None
    alarm_xy: tuple[float, float] | None = None
    job_id: str = ""
    project_title: str = "Remote 30 전체 배관망 총괄"


# ────────────────────────────────────────────────────────────────────────────
# RiserTables — Stage B 산출 / Stage C 입력
# ────────────────────────────────────────────────────────────────────────────

@dataclass
class RiserTables:
    """라이저(펌프~AV) 구간의 PipeTables 확장.

    prototype.PipeTables 와 호환되는 nodes/pipes 형식 + pumps/valves 추가.
    각 dict 의 키는 emit_sdf 의 직렬화에 그대로 쓰일 수 있도록 PIPENET 속성명에 맞춤.
    """
    nodes: list[dict] = field(default_factory=list)
    pipes: list[dict] = field(default_factory=list)
    pumps: list[dict] = field(default_factory=list)   # Pump-fan 직렬화용
    valves: list[dict] = field(default_factory=list)  # Elastomeric-valve 직렬화용
    av_node_label: str = ""  # 라이저 끝점 = 헤드망 source 와 stitch 할 노드 라벨


# ────────────────────────────────────────────────────────────────────────────
# Stage A — 헤드망 추출 (remote30_prototype 재사용)
# ────────────────────────────────────────────────────────────────────────────

def run_stage_a(inputs: OverallInputs) -> Iterator[dict]:
    """Stage A — 평면도 DXF → 헤드 30개 + 배관망 추출 이벤트 스트림.

    호출 측(서버)은 마지막 ``stage2_complete`` 이벤트의 데이터를 보관 후
    사용자 헤드 편집 단계 → run_stages_3_5 → build_input_tables 까지 진행한다.
    """
    from remote30_prototype import run_stages_0_2
    yield from run_stages_0_2(
        dxf_path=inputs.dxf_path,
        job_id=inputs.job_id,
        alarm_xy=inputs.alarm_xy,
    )


# ────────────────────────────────────────────────────────────────────────────
# Stage B — zone별 라이저 템플릿
# ────────────────────────────────────────────────────────────────────────────
#
# 좌표 처리 방침:
# PIPENET SDF 의 Position x/y 는 isometric 시각 표시용일 뿐 수리계산엔 무관하다.
# emit_sdf 의 _xform 이 모든 노드 좌표를 bbox 중심 (0,0) + 약 3000 unit 으로 정규화하므로
# 라이저 노드 좌표의 절댓값은 의미가 없고 상대 배치(라이저 수직 진행, AV 가 적절한 위치)
# 만 유지하면 된다. 그러므로 빌딩 무관 logical 좌표계로 정의 — 모든 빌딩에서 동작.
#
# Logical 좌표 (mm 단위, 옥상 수원을 (0, 0) 으로):
#   Input(1)            (   0,    0)    옥상 수원
#   N2                  (-500, -100)    옥상 수평 분기
#   N3                  (-500, -300)    옥상 → 라이저 진입 (수직 강하 시작)
#   N4                  (-300, -1500)   라이저 중간 (수평 우회)
#   N7                  (-300, -2800)   PRV 진입 직전
#   N8 (PRV in)         (-200, -2900)
#   N89 (PRV out)       (-100, -2950)
#   N5 (AV 직전)         (-300, -3300)
#   N10 (AV ★)          ( 100, -3300)   헤드망 source 와 stitch
#   N87 (2차 PRV in)    ( -50, -3100)   LLSP_2STAGE 만
#   N88 (2차 PRV out)   (  50, -3150)
#   N100 (Pump Input)   ( 200,   100)   HSP_PUMP 만 (Input 1 보다 더 위)
#
# v1 구현 (대명동 201동 답안 SDF) 의 좌표는 git history 에서 확인 가능.
#
# 라이저 노드 라벨 컨벤션 (답안과 정합)::
#     1   : Input 노드 (옥상 수원, 1기압 boundary)
#     2,3,4,7  : 라이저 중간 노드 (옥상→하강→PRV 직전)
#     8   : 1차 PRV in
#     89  : 1차 PRV out
#     5   : AV 직전 노드
#     10  : AV 노드 (헤드망 source 와 stitch)
#     87  : 2차 PRV in (LLSP_2STAGE 만)
#     88  : 2차 PRV out (LLSP_2STAGE 만)
#     100 : Pump-fan 전 Input (HSP_PUMP 만 — Input 라벨이 1 → 100 으로 옮김)
#
# ────────────────────────────────────────────────────────────────────────────


def _node(label: str, elev: float, x: float, y: float, *,
          io_node: str = "No", pressure_pa: float | None = None) -> dict:
    """라이저 노드 dict — remote30_prototype.PipeTables.nodes 호환 형식."""
    d: dict = {
        "label": label, "elevation": elev, "io_node": io_node,
        "x": int(round(x)), "y": int(round(y)),
    }
    if pressure_pa is not None:
        d["pressure_pa"] = pressure_pa  # Input node 의 <Calculation-spec pressure="..."/>
    return d


def _pipe(label: str, in_lbl: str, out_lbl: str, bore_mm: int,
          length_m: float, rise_m: float = 0.0, c_factor: str = "120") -> dict:
    """라이저 파이프 dict — remote30_prototype.PipeTables.pipes 호환."""
    return {
        "label": label, "in": in_lbl, "out": out_lbl, "type": "KSD 3507",
        "dia": bore_mm, "length": round(length_m, 2), "elev": rise_m,
        "c": c_factor, "status": "Normal", "group": "Unset",
    }


def _pump_fan(label: str, in_lbl: str, out_lbl: str, *,
              library_pump: str, efficiency: int = 100, status: int = 1) -> dict:
    """Pump-fan dict — emit_full_sdf 가 <Pump-fan> 으로 직렬화."""
    return {
        "label": label, "in": in_lbl, "out": out_lbl,
        "efficiency": efficiency, "status": status,
        "library_pump": library_pump,
        "percentage_open": 1,
    }


def _pressure_valve(label: str, in_lbl: str, out_lbl: str, *,
                    target_pa: float, valve_type: str = "output") -> dict:
    """Elastomeric-valve dict — emit_full_sdf 가 <Elastomeric-valve> 로 직렬화."""
    return {
        "label": label, "in": in_lbl, "out": out_lbl,
        "target_value": float(target_pa), "type": valve_type,
    }


def _elev_at_floor(profile: BuildingPressureProfile | None, floor_label: str,
                   fallback: float = 0.0) -> float:
    """압력표에서 층 라벨의 누적 낙차 (m) → SDF elevation (음수)."""
    if profile is not None:
        row = profile.find_by_label(floor_label)
        if row is not None:
            return -float(row.head_drop_m)
    return fallback


# ── 라이저 logical 좌표 — 빌딩 무관. emit_sdf 의 _xform 이 정규화.
_COORDS_INPUT      = (   0,     0)    # 옥상 수원
_COORDS_N2         = (-500,  -100)    # 옥상 수평 분기
_COORDS_N3         = (-500,  -300)    # 옥상 → 라이저 진입
_COORDS_N4         = (-300, -1500)    # 라이저 중간 (수평 우회)
_COORDS_N7         = (-300, -2800)    # PRV 진입 직전
_COORDS_PRV_IN     = (-200, -2900)    # 노드 8 (1차 PRV in)
_COORDS_PRV_OUT    = (-100, -2950)    # 노드 89 (1차 PRV out)
_COORDS_PRV2_IN    = ( -50, -3100)    # 노드 87 (2차 PRV in, LLSP)
_COORDS_PRV2_OUT   = (  50, -3150)    # 노드 88 (2차 PRV out, LLSP)
_COORDS_N5_AV_PREV = (-300, -3300)    # AV 직전
_COORDS_AV         = ( 100, -3300)    # 노드 10 (헤드망 source ★)
_COORDS_PUMP_INPUT = ( 200,   100)    # HSP Pump-fan 앞 Input(100), Input(1) 보다 위


def build_riser_lsp_1stage(spec: ZoneSpec, profile: BuildingPressureProfile | None) -> RiserTables:
    """LSP 1차감압 라이저 — 자연낙차 + PRV 1개.

    노드: Input(1, elev=0) → 2 → 3 → 4 → 7 → 8(PRV in) → 89(PRV out) → 5 → 10(AV)
    파이프: 1→2, 2→3, 3→4, 4→7, 7→8, 89→5, 5→10
    Elastomeric-valve: 8 → 89 (target = spec.prv1_target_pa)
    """
    if spec.prv1_target_pa is None:
        raise ValueError("LSP_1STAGE 는 prv1_target_pa 가 필요합니다 (kg/cm² 또는 m 수두 → Pa).")
    elev_av = _elev_at_floor(profile, spec.target_floor, fallback=-98.45)
    elev_prv = -73.35  # PRV 위치 (옥상 -73m) — 답안 16F 기준 고정값 (24F 1차 감압 위치)
    elev_top = -3.75   # 옥상 수직 분기 위치
    return RiserTables(
        nodes=[
            _node("1",  0.0,        *_COORDS_INPUT,    io_node="Input", pressure_pa=ATM_PA),
            _node("2",  0.0,        *_COORDS_N2),
            _node("3",  elev_top,   *_COORDS_N3),
            _node("4",  elev_top,   *_COORDS_N4),
            _node("7",  elev_prv,   *_COORDS_N7),
            _node("8",  elev_prv,   *_COORDS_PRV_IN),
            _node("89", elev_prv,   *_COORDS_PRV_OUT),
            _node("5",  elev_av - 1.0, *_COORDS_N5_AV_PREV),
            _node("10", elev_av,    *_COORDS_AV),
        ],
        pipes=[
            _pipe("r1", "1", "2",   150, 20.95, 0.0),
            _pipe("r2", "2", "3",   150,  3.75, elev_top - 0.0),
            _pipe("r3", "3", "4",   150, 14.93, 0.0),
            _pipe("r4", "4", "7",   150, abs(elev_prv - elev_top), elev_prv - elev_top),
            _pipe("r5", "7", "8",   150,  0.5,  0.0),
            _pipe("r6", "89","5",   150, abs(elev_av - 1.0 - elev_prv), (elev_av - 1.0) - elev_prv),
            _pipe("r7", "5", "10",  125,  1.5,  1.0),
        ],
        pumps=[],
        valves=[
            _pressure_valve("1", "8", "89", target_pa=spec.prv1_target_pa),
        ],
        av_node_label="10",
    )


def build_riser_hsp_pump(spec: ZoneSpec, profile: BuildingPressureProfile | None) -> RiserTables:
    """HSP 펌프식 라이저 — 자연낙차 부족 → Pump-fan 부스터 + PRV 1개.

    노드: Input(100, elev=0) → [Pump-fan 100→1] → 1 → 2 → 3 → 4 → 7 → 8(PRV in) → 89(PRV out) → 5 → 10(AV)
    Pump-fan: 100 → 1 (Library-pump = spec.pump_library_name, count=spec.pump_count)
    """
    if spec.prv1_target_pa is None:
        raise ValueError("HSP_PUMP 는 prv1_target_pa 가 필요합니다.")
    # HSP 는 고층부라 elev_av 가 양수 (옥상보다 위)가 아니라 옥상 근처. 답안 29F = -64.2m 정도
    elev_av = _elev_at_floor(profile, spec.target_floor, fallback=-67.1)
    elev_prv = -10.0   # HSP 1차 PRV 위치 (옥상 직하단 보통)
    elev_top = -3.75
    nodes = [
        _node("100", 0.0, *_COORDS_PUMP_INPUT, io_node="Input", pressure_pa=ATM_PA),
        _node("1",   0.0,        *_COORDS_INPUT),
        _node("2",   0.0,        *_COORDS_N2),
        _node("3",   elev_top,   *_COORDS_N3),
        _node("4",   elev_top,   *_COORDS_N4),
        _node("7",   elev_prv,   *_COORDS_N7),
        _node("8",   elev_prv,   *_COORDS_PRV_IN),
        _node("89",  elev_prv,   *_COORDS_PRV_OUT),
        _node("5",   elev_av - 1.0, *_COORDS_N5_AV_PREV),
        _node("10",  elev_av,    *_COORDS_AV),
    ]
    pipes = [
        _pipe("r1", "1", "2",   150, 20.95, 0.0),
        _pipe("r2", "2", "3",   150,  3.75, elev_top),
        _pipe("r3", "3", "4",   150, 14.93, 0.0),
        _pipe("r4", "4", "7",   150, abs(elev_prv - elev_top), elev_prv - elev_top),
        _pipe("r5", "7", "8",   150,  0.5,  0.0),
        _pipe("r6", "89","5",   150, abs(elev_av - 1.0 - elev_prv), (elev_av - 1.0) - elev_prv),
        _pipe("r7", "5", "10",  125,  1.5,  1.0),
    ]
    pumps = [
        _pump_fan(str(i + 1), "100" if i == 0 else "100", "1",
                  library_pump=spec.pump_library_name)
        for i in range(max(1, spec.pump_count))
    ]
    return RiserTables(
        nodes=nodes, pipes=pipes, pumps=pumps,
        valves=[_pressure_valve("1", "8", "89", target_pa=spec.prv1_target_pa)],
        av_node_label="10",
    )


def build_riser_llsp_2stage(spec: ZoneSpec, profile: BuildingPressureProfile | None) -> RiserTables:
    """LLSP 2차감압 라이저 — 자연낙차 + 1차 PRV + 2차 PRV (지하주차장).

    노드: ... 89(1차 PRV out) → 87(2차 PRV in) → 88(2차 PRV out) → 5 → 10(AV)
    Elastomeric-valve: 8→89 (1차, target=prv1_target_pa), 87→88 (2차, target=prv2_target_pa)
    """
    if spec.prv1_target_pa is None or spec.prv2_target_pa is None:
        raise ValueError("LLSP_2STAGE 는 prv1_target_pa, prv2_target_pa 모두 필요합니다.")
    elev_av = _elev_at_floor(profile, spec.target_floor, fallback=-159.6)  # B1 = -159.6m
    elev_prv1 = -73.35   # 1차 PRV (옥상 -73m)
    elev_prv2 = -145.4   # 2차 PRV (1.5F, 옥상 -145m)
    elev_top = -3.75
    nodes = [
        _node("1",  0.0,         *_COORDS_INPUT,    io_node="Input", pressure_pa=ATM_PA),
        _node("2",  0.0,         *_COORDS_N2),
        _node("3",  elev_top,    *_COORDS_N3),
        _node("4",  elev_top,    *_COORDS_N4),
        _node("7",  elev_prv1,   *_COORDS_N7),
        _node("8",  elev_prv1,   *_COORDS_PRV_IN),
        _node("89", elev_prv1,   *_COORDS_PRV_OUT),
        _node("87", elev_prv2,   *_COORDS_PRV2_IN),
        _node("88", elev_prv2,   *_COORDS_PRV2_OUT),
        _node("5",  elev_av - 1.0, *_COORDS_N5_AV_PREV),
        _node("10", elev_av,     *_COORDS_AV),
    ]
    pipes = [
        _pipe("r1", "1", "2",   150, 20.95, 0.0),
        _pipe("r2", "2", "3",   150,  3.75, elev_top),
        _pipe("r3", "3", "4",   150, 14.93, 0.0),
        _pipe("r4", "4", "7",   150, abs(elev_prv1 - elev_top), elev_prv1 - elev_top),
        _pipe("r5", "7", "8",   150,  0.5,  0.0),
        _pipe("r6", "89", "87", 150, abs(elev_prv2 - elev_prv1), elev_prv2 - elev_prv1),
        _pipe("r7", "88", "5",  150, abs(elev_av - 1.0 - elev_prv2), (elev_av - 1.0) - elev_prv2),
        _pipe("r8", "5", "10",  125,  1.5,  1.0),
    ]
    return RiserTables(
        nodes=nodes, pipes=pipes, pumps=[],
        valves=[
            _pressure_valve("1", "8", "89", target_pa=spec.prv1_target_pa),
            _pressure_valve("2", "87", "88", target_pa=spec.prv2_target_pa),
        ],
        av_node_label="10",
    )


def build_riser(spec: ZoneSpec, profile: BuildingPressureProfile | None) -> RiserTables:
    """zone 분기 라우터."""
    if spec.zone_type == ZoneType.HSP_PUMP:
        return build_riser_hsp_pump(spec, profile)
    if spec.zone_type == ZoneType.LSP_1STAGE:
        return build_riser_lsp_1stage(spec, profile)
    if spec.zone_type == ZoneType.LLSP_2STAGE:
        return build_riser_llsp_2stage(spec, profile)
    if spec.zone_type == ZoneType.LSP_GRAVITY:
        # 자연낙차 감압 없음 — LSP_1STAGE 의 PRV 만 제거한 변형
        rt = build_riser_lsp_1stage(
            ZoneSpec(zone_type=ZoneType.LSP_1STAGE, target_floor=spec.target_floor,
                     prv1_target_pa=ATM_PA),  # dummy — valves 비울 거니까 무시
            profile,
        )
        rt.valves = []
        # PRV 자리 노드 8, 89 도 제거하고 7 → 5 직결
        rt.nodes = [n for n in rt.nodes if n["label"] not in ("8", "89")]
        rt.pipes = [p for p in rt.pipes if p["in"] not in ("7", "89") or p["out"] not in ("8", "5")]
        rt.pipes.append(_pipe("r_g", "7", "5", 150, 30.0, -30.0))  # 7 → 5 직결 단순화
        return rt
    raise ValueError(f"Unknown zone_type: {spec.zone_type}")


# ────────────────────────────────────────────────────────────────────────────
# Stage C — 라이저 ↔ 헤드망 stitch
# ────────────────────────────────────────────────────────────────────────────


@dataclass
class CombinedTables:
    """라이저 + 헤드망 결합 결과 — emit_full_sdf 입력."""
    nodes: list[dict] = field(default_factory=list)
    pipes: list[dict] = field(default_factory=list)
    nozzles: list[dict] = field(default_factory=list)
    fittings: list[dict] = field(default_factory=list)
    equipment: list[dict] = field(default_factory=list)
    pumps: list[dict] = field(default_factory=list)
    valves: list[dict] = field(default_factory=list)
    meta: list[tuple[str, str]] = field(default_factory=list)


def stitch_riser_and_heads(riser: RiserTables, head_tables: Any) -> CombinedTables:
    """라이저 끝점(AV node) ↔ 헤드망 source(=같은 label) 결합.

    Args:
        riser: Stage B 산출.
        head_tables: remote30_prototype.PipeTables 인스턴스.

    Returns:
        CombinedTables — 모든 element 가 합쳐진 표.

    충돌 처리:
        라이저 노드 라벨 = {1..9, 87, 88, 89, 100} + AV(10)
        헤드망 노드 라벨 = {10, 11, 12, ...} (10 이 source = AV)
        AV(10) 만 공통 — 라이저 쪽 노드만 유지, 헤드망 쪽 노드 10 의 elevation
        을 라이저 AV elevation 으로 동기화.
    """
    av_lbl = riser.av_node_label
    riser_av_node = next((n for n in riser.nodes if n["label"] == av_lbl), None)
    if riser_av_node is None:
        raise ValueError(f"라이저에 AV 노드(label={av_lbl})가 없음")

    # 헤드망 노드 10 의 elevation → 라이저 AV elevation 으로 일치
    head_nodes_filtered = []
    for n in head_tables.nodes:
        if n["label"] == av_lbl:
            # AV 는 라이저 쪽에서 이미 포함 — 헤드망 쪽 사본 skip
            continue
        head_nodes_filtered.append(n)

    return CombinedTables(
        nodes=list(riser.nodes) + head_nodes_filtered,
        pipes=list(riser.pipes) + list(head_tables.pipes),
        nozzles=list(head_tables.nozzles),
        fittings=list(head_tables.fittings),
        equipment=list(head_tables.equipment),
        pumps=list(riser.pumps),
        valves=list(riser.valves),
        meta=list(getattr(head_tables, "meta", [])),
    )


# ────────────────────────────────────────────────────────────────────────────
# Stage D — emit_full_sdf (PIPENET-native 후처리 + Pump-fan / Elastomeric-valve)
# ────────────────────────────────────────────────────────────────────────────


def emit_full_sdf(combined: CombinedTables, out_path: Path, *,
                  project_title: str = "Remote 30 전체 배관망 총괄") -> Path:
    """완성 SDF 직렬화.

    1단계: ``remote30_prototype.emit_sdf`` 호출 — PIPENET-native 후처리
           (빈 Pipe-set placeholder, 6 schedule embed, SLF 동봉, Template Graphics)
           가 모두 그대로 적용됨.
    2단계: 결과 SDF 를 다시 열어:
           - Input 노드에 ``<Calculation-spec pressure="ATM_PA"/>`` 추가
           - ``<Pump-fan>`` element 추가 (combined.pumps)
           - ``<Elastomeric-valve>`` element 추가 (combined.valves)
    """
    from remote30_prototype import emit_sdf, PipeTables

    # 1단계: PipeTables 로 캐스팅 후 emit_sdf 호출
    tables = PipeTables(
        nodes=list(combined.nodes),
        pipes=list(combined.pipes),
        nozzles=list(combined.nozzles),
        fittings=list(combined.fittings),
        equipment=list(combined.equipment),
        meta=list(combined.meta),
    )
    emit_sdf(tables, out_path, project_title=project_title)

    # 2단계: SDF 재오픈 → Pump-fan / Elastomeric-valve / Calculation-spec 추가
    import xml.etree.ElementTree as ET
    tree = ET.parse(out_path)
    root = tree.getroot()

    # Input 노드에 boundary pressure 명시 (Calculation-spec)
    pressure_by_label: dict[str, float] = {
        n["label"]: float(n["pressure_pa"])
        for n in combined.nodes if n.get("pressure_pa") is not None
    }
    for node_el in root.iter("Node"):
        lbl = node_el.get("label", "")
        if lbl in pressure_by_label and node_el.find("Calculation-spec") is None:
            ET.SubElement(node_el, "Calculation-spec",
                          {"pressure": str(int(pressure_by_label[lbl]))})

    # Links 안에 Pump-fan + Elastomeric-valve 삽입 (Nozzle 다음, 또는 끝쪽)
    for links in root.iter("Links"):
        for pump in combined.pumps:
            pf = ET.Element("Pump-fan", {
                "efficiency": str(pump["efficiency"]),
                "input": pump["in"],
                "label": pump["label"],
                "output": pump["out"],
                "status": str(pump["status"]),
            })
            ET.SubElement(pf, "Description")
            lib = ET.SubElement(pf, "Library-pump")
            lib.text = pump["library_pump"]
            ET.SubElement(pf, "Pump-setting",
                          {"percentage-open": str(pump["percentage_open"])})
            links.append(pf)

        for v in combined.valves:
            ev = ET.Element("Elastomeric-valve", {
                "input": v["in"],
                "label": v["label"],
                "output": v["out"],
                "target-value": f"{v['target_value']:.6g}",
                "type": v["type"],
            })
            links.append(ev)
        break  # 첫 Links 만

    tree.write(out_path, encoding="utf-8", xml_declaration=True)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# 통합 파이프라인 — 한 번에 모든 Stage 진행 (서버 API 에서 호출)
# ────────────────────────────────────────────────────────────────────────────


def run_full_pipeline(inputs: OverallInputs, head_tables: Any, out_path: Path) -> Path:
    """Stage B + C + D 일괄 실행. Stage A 는 별도 (서버 측에서 사용자 헤드 편집 단계 필요).

    Args:
        inputs: OverallInputs (zone_spec, profile, project_title 사용).
        head_tables: remote30_prototype.PipeTables — Stage A 의 결과.
        out_path: 출력 SDF 경로.
    """
    riser = build_riser(inputs.zone_spec, inputs.profile)
    combined = stitch_riser_and_heads(riser, head_tables)
    return emit_full_sdf(combined, out_path, project_title=inputs.project_title)


# ────────────────────────────────────────────────────────────────────────────
# 직접 입력 폼 → ZoneSpec / BuildingPressureProfile 변환
# ────────────────────────────────────────────────────────────────────────────

def zone_spec_from_form(form: dict[str, Any]) -> ZoneSpec:
    """HTML 폼 데이터 → ZoneSpec.

    필수 폼 필드::

        zone_type           : "hsp_pump" / "lsp_1stage" / "llsp_2stage" / "lsp_gravity"
        target_floor        : "16층" 등 라벨

    선택 폼 필드 (감압 zone)::

        prv1_target_kgf     : 1차 PRV 출구압 (kg/cm²) — Pa 변환됨
        prv2_target_kgf     : 2차 PRV 출구압 (kg/cm²) — LLSP_2STAGE 만
        prv1_target_m       : 1차 PRV 출구압 (m 수두) — kg/cm² 와 둘 중 하나만
        prv2_target_m       : 2차 PRV 출구압 (m 수두)
        pump_library_name   : Library-pump 이름 (HSP_PUMP, 기본 SP_162M_2900LPM)
        pump_count          : Pump-fan 개수 (HSP_PUMP, 기본 2)
    """
    zone_type = ZoneType(form.get("zone_type", "lsp_1stage"))

    def _to_pa(kgf_key: str, m_key: str) -> float | None:
        kgf = form.get(kgf_key, "")
        m = form.get(m_key, "")
        if kgf:
            try:
                return float(kgf) * KGF_CM2_TO_PA
            except ValueError:
                pass
        if m:
            try:
                return float(m) * M_TO_PA
            except ValueError:
                pass
        return None

    return ZoneSpec(
        zone_type=zone_type,
        target_floor=str(form.get("target_floor", "")).strip(),
        prv1_target_pa=_to_pa("prv1_target_kgf", "prv1_target_m"),
        prv2_target_pa=_to_pa("prv2_target_kgf", "prv2_target_m"),
        pump_library_name=str(form.get("pump_library_name", "SP_162M_2900LPM")).strip(),
        pump_count=int(form.get("pump_count", 2)),
    )


def parse_system_diagram_dxf(dxf_path: Path, *, default_height_m: float = 2.9,
                              roof_height_m: float = 6.0) -> BuildingPressureProfile:
    """계통도 DXF 의 텍스트 라벨 → BuildingPressureProfile 자동 추정.

    추출 패턴 (대명동 201동 계통도 분석 기반):
        "지상N층"   → "N층" label
        "지하N층"   → "지하N층" label
        "옥상", "PH", "PH N F"  → "옥상층" label
        "Nf" / "NF" / "NFL"  (fallback)

    elevation 계산:
        TEXT 의 Y 좌표를 정렬 (큰 Y = 위쪽) → 최상부부터 default_height_m 씩 누적.
        도면의 Y 단위가 도면별로 달라 직접 사용은 불안정. 층고는 default 로 채우고
        사용자가 검토/수정.

    Args:
        dxf_path: 계통도 DXF.
        default_height_m: 표준 층고 (기본 2.9m).
        roof_height_m: 옥상층의 층고 (기본 6m — PDF 압력표 기준).

    Returns:
        BuildingPressureProfile — 최상부 → 최하부 순. 비어있을 수도 (라벨 0건).
    """
    import re
    try:
        import ezdxf as _ezdxf
    except ImportError as exc:
        raise RuntimeError("ezdxf 가 설치되지 않아 계통도 DXF 파싱 불가") from exc

    doc = _ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()

    PAT_ABOVE = re.compile(r"지상\s*(\d+)\s*층")
    PAT_BELOW = re.compile(r"지하\s*(\d+)\s*층")
    PAT_ROOF  = re.compile(r"옥상|PH(?:\s*\d+)?\s*F?")
    PAT_NF    = re.compile(r"^\s*(\d{1,2})\s*F\s*$")

    candidates: list[tuple[str, float]] = []
    for e in msp:
        if e.dxftype() not in ("TEXT", "MTEXT"):
            continue
        try:
            v = e.dxf.text if e.dxftype() == "TEXT" else e.text
        except Exception:
            continue
        if not v or not v.strip():
            continue
        v = v.strip()
        try:
            pos = e.dxf.insert
            y = float(pos.y)
        except Exception:
            continue
        if (m := PAT_ABOVE.search(v)):
            candidates.append((f"{m.group(1)}층", y))
            continue
        if (m := PAT_BELOW.search(v)):
            candidates.append((f"지하{m.group(1)}층", y))
            continue
        if (m := PAT_NF.search(v)):
            candidates.append((f"{m.group(1)}층", y))
            continue
        if PAT_ROOF.search(v):
            candidates.append(("옥상층", y))

    # 같은 label 중복 시 가장 큰 Y 사용 (도면 상단 라벨이 가장 신뢰)
    by_label: dict[str, float] = {}
    for label, y in candidates:
        if label not in by_label or y > by_label[label]:
            by_label[label] = y

    # Y 내림차순 정렬 — 최상부 (Y 큰 것) → 최하부
    sorted_floors = sorted(by_label.items(), key=lambda kv: -kv[1])

    rows: list[FloorRow] = []
    cumulative_drop = 0.0
    for i, (label, _y) in enumerate(sorted_floors):
        height = roof_height_m if label == "옥상층" else default_height_m
        if i > 0:
            cumulative_drop += height
        rows.append(FloorRow(
            floor_label=label,
            height_m=height,
            head_drop_m=round(cumulative_drop, 1),
            note="(계통도 자동 추출 — 층고/낙차 검토 필요)",
        ))

    return BuildingPressureProfile(
        building_name=dxf_path.stem,
        floors=rows,
    )


def profile_from_form(form: dict[str, Any]) -> BuildingPressureProfile | None:
    """HTML 폼의 row 배열(JSON) → BuildingPressureProfile.

    폼 필드 ``pressure_table_json`` 이 있으면 파싱, 없으면 None.
    JSON 형식 예::

        [{"floor_label": "옥상층", "height_m": 6,   "head_drop_m": 3.1},
         {"floor_label": "49층",   "height_m": 3.1, "head_drop_m": 6.2},
         ...]
    """
    raw = form.get("pressure_table_json", "").strip()
    if not raw:
        return None
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return None
    rows = [FloorRow(
        floor_label=str(d.get("floor_label", "")).strip(),
        height_m=float(d.get("height_m", 0) or 0),
        head_drop_m=float(d.get("head_drop_m", 0) or 0),
        after_prv_m=(float(d["after_prv_m"]) if d.get("after_prv_m") not in (None, "") else None),
        note=str(d.get("note", "") or ""),
    ) for d in data if d.get("floor_label")]
    return BuildingPressureProfile(
        building_name=str(form.get("building_name", "")).strip(),
        floors=rows,
    )

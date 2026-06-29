"""HAS (HASS / 하스) ``.has`` 배관망 파일 변환기 — CommonNetwork ↔ .has.

HASS(국산 소방 수리계산 프로그램) ``.has`` = 단일 UTF-8 JSON(``Document`` 루트).
PIPENET ``.sdf``(XML 분리형)·K-Fire Solver ``.kfp``(JSON 입력/결과 분리)와 달리
**도면 + 입력 + 결과(A/B 2시나리오)를 한 파일에 통합**한다. 본 모듈은
``kfp_sdf_converter`` 의 ``CommonNetwork`` 를 재사용해 SDF/KFP ↔ HAS 를 잇는다.

설계 노트 (has-format-notes 메모리 + 붙임 샘플 실측):
- 노드 식별 = ``nodeBlks[].Id`` (int). ``plines``/``nozzleBlks``/``pumpBlks`` 가
  ``SNodeId``/``ENodeId`` 로 이 Id 를 참조(Label 은 표시용).
- ``IoNode`` 1 = 시스템 source(경계), 0 = 일반. ``Height`` = elevation(m).
  ``InsertionPoint``/``Points`` = ``"x, y, z"`` 문자열(m 단위).
- nozzle K 는 노즐 블록에 직접 없고 ``PipeDbMgr.nozzleDataTable[NozzleName==NozzleType].KFactor``
  (SI = L/min·bar^-0.5) 로 lookup.
- pline fitting = ``Cnt*`` **개수 카운트**(elbow/tee/valve…). ``RESULT_A/B_*`` 는
  HASS 자체 계산 → emit 시 "0" 으로 초기화해 HASS 가 재계산.
- emit 은 붙임 샘플을 **스켈레톤**으로 로드(참조 DB·범례·설정·단위계) 후
  ``PipeInfoMgr`` 의 geometry 블록만 교체. 처음부터 합성하면 HASS 거부 위험.

한계 (v1):
- PRV(reducingBlks)·orifice(orifiaceBlks)·장식용 밸브 심볼 블록은 parse 는 하되
  emit 은 하지 않음(우리 통합망에 없음). 밸브의 수리 영향은 pline ``Cnt*`` 로 반영.
- 배관 재질은 자동분류 금지(pipe-material-not-in-dxf 메모리) → 스켈레톤 기본값 고정.
"""

from __future__ import annotations

import copy
import json
import math
import os
from pathlib import Path

from kfp_sdf_converter import (
    CommonFitting,
    CommonNetwork,
    CommonNode,
    CommonPipe,
    _BAR_TO_M,
    _coord_scale,
    _spread_factor,
    emit_sdf_xml,
    parse_sdf,
    simplify_passthrough_nodes,
)

# ────────────────────────────────────────────────────────────────────────────
# 스켈레톤 템플릿 — 붙임 샘플(.has) 재사용
# ────────────────────────────────────────────────────────────────────────────

_MOD_ROOT = Path(__file__).resolve().parent


def _template_path() -> Path:
    """HAS emit 스켈레톤(.has) 경로. env override → scripts/ → 모듈 루트 순.

    REMOTE30_HAS_TEMPLATE 로 명시 가능. 없으면 scripts/ 의 붙임 샘플을 사용한다
    (lh-airgap 메모리: sprinkler_ai_agent_server_source_* 폴더는 LH 패키지 제외
    대상이라 scripts/ 사본을 우선).
    """
    env = os.environ.get("REMOTE30_HAS_TEMPLATE")
    if env and Path(env).is_file():
        return Path(env)
    cands = [
        _MOD_ROOT / "scripts" / "붙임1. 102동 25층 수리계산서.has",
        *sorted(_MOD_ROOT.glob("scripts/*.has")),
        *sorted(_MOD_ROOT.glob("*.has")),
    ]
    for c in cands:
        if c.is_file():
            return c
    raise FileNotFoundError(
        "HAS 스켈레톤 템플릿(.has)을 찾을 수 없음 — REMOTE30_HAS_TEMPLATE 환경변수로 "
        f"지정하거나 '{_MOD_ROOT / 'scripts'}' 에 .has 샘플을 두세요."
    )


# ────────────────────────────────────────────────────────────────────────────
# fitting ↔ Cnt* 매핑 (CommonFitting.type_id ↔ HAS pline 카운트 필드)
# ────────────────────────────────────────────────────────────────────────────

# CommonFitting.type_id (KFP/SDF 정규화: 소문자 + '-' 구분) → HAS Cnt* 필드.
_FITTING_TO_CNT: dict[str, str] = {
    "elbow": "Cnt90DegreeElbow",
    "elbow-90": "Cnt90DegreeElbow",
    "90-elbow": "Cnt90DegreeElbow",
    "elbow-standard": "Cnt90DgreeStandardElbow",
    "standard-elbow": "Cnt90DgreeStandardElbow",
    "elbow-45": "Cnt45DegreeElbow",
    "45-elbow": "Cnt45DegreeElbow",
    "tee": "CntDivideTee",
    "divide-tee": "CntDivideTee",
    "tee-divide": "CntDivideTee",
    "tee-branch": "CntDivideTee",
    "gate": "CntGateValve",
    "gate-valve": "CntGateValve",
    "check": "CntSwingcheckValve",
    "check-valve": "CntSwingcheckValve",
    "swing-check": "CntSwingcheckValve",
    "swingcheck": "CntSwingcheckValve",
    "alarm": "CntAlarmValve",
    "alarm-valve": "CntAlarmValve",
    "butterfly": "CntButterflyValve",
    "butterfly-valve": "CntButterflyValve",
    "angle": "CntAngleValve",
    "angle-valve": "CntAngleValve",
    "strainer": "CntStrainer",
    "flexible": "CntFlexibleJoint",
    "flexible-joint": "CntFlexibleJoint",
    "deluge": "CntDelugeValve",
    "deluge-valve": "CntDelugeValve",
    "dry": "CntDryValve",
    "dry-valve": "CntDryValve",
    "preaction": "CntPreactionValve",
    "preaction-valve": "CntPreactionValve",
}

# HAS pline 의 전체 Cnt* 필드 (emit 시 미사용분은 "0").
_CNT_FIELDS: tuple[str, ...] = (
    "Cnt45DegreeElbow",
    "Cnt90DegreeElbow",
    "Cnt90DgreeStandardElbow",
    "CntAlarmValve",
    "CntAngleValve",
    "CntButterflyValve",
    "CntDelugeValve",
    "CntDivideTee",
    "CntDryValve",
    "CntEtcPartLengthValve",
    "CntFlexibleJoint",
    "CntGateValve",
    "CntPreactionValve",
    "CntStrainer",
    "CntSwingcheckValve",
)

# 역매핑 (parse): Cnt* → 대표 CommonFitting.type_id.
_CNT_TO_FITTING: dict[str, str] = {
    "Cnt90DegreeElbow": "elbow",
    "Cnt90DgreeStandardElbow": "elbow-standard",
    "Cnt45DegreeElbow": "elbow-45",
    "CntDivideTee": "tee",
    "CntGateValve": "gate-valve",
    "CntSwingcheckValve": "check",
    "CntAlarmValve": "alarm-valve",
    "CntButterflyValve": "butterfly",
    "CntAngleValve": "angle",
    "CntStrainer": "strainer",
    "CntFlexibleJoint": "flexible-joint",
    "CntDelugeValve": "deluge",
    "CntDryValve": "dry",
    "CntPreactionValve": "preaction",
}


# ────────────────────────────────────────────────────────────────────────────
# 좌표 유틸
# ────────────────────────────────────────────────────────────────────────────


def _xyz_from_str(s: str) -> tuple[float, float, float]:
    """HAS 좌표 문자열 ``"x, y, z"`` → (x, y, z). 빈 값은 0."""
    parts = [p.strip() for p in str(s).split(",")]

    def _f(i: int) -> float:
        if i < len(parts) and parts[i]:
            try:
                return float(parts[i])
            except ValueError:
                return 0.0
        return 0.0

    return _f(0), _f(1), _f(2)


def _xyz_to_str(x: float, y: float, z: float = 0.0) -> str:
    return f"{x:.6f}, {y:.6f}, {z:.6f}"


# 좌표 표시배율 헬퍼(_coord_scale / _spread_factor)는 kfp_sdf_converter 에서 import.
# KFP·HAS 가 동일 통합망에서 동일 배율을 쓰도록 단일 소스 유지(중복 제거).


# ────────────────────────────────────────────────────────────────────────────
# parse_has — .has(JSON) → CommonNetwork
# ────────────────────────────────────────────────────────────────────────────


def parse_has(path: str | Path) -> CommonNetwork:
    """HAS 파일(.has, UTF-8 JSON) → CommonNetwork."""
    data = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    return has_dict_to_network(data)


def has_dict_to_network(data: dict) -> CommonNetwork:
    """HAS dict → CommonNetwork.

    nodeBlks → CommonNode(base), nozzleBlks/pumpBlks/reducingBlks/orifiaceBlks 가
    참조하는 노드의 kind 를 승격(nozzle/pump/prv/orifice). plines → CommonPipe
    (Cnt* → CommonFitting). 노드 식별은 정수 ``Id`` 문자열.
    """
    doc = data.get("Document", data)
    pim = doc.get("PipeInfoMgr", {}) or {}
    db = doc.get("PipeDbMgr", {}) or {}
    net = CommonNetwork(source_format="has")

    # nozzle K 룩업 테이블 (NozzleName → row)
    ktab: dict[str, dict] = {}
    for row in db.get("nozzleDataTable", []) or []:
        ktab[str(row.get("NozzleName"))] = row

    # ── 노드
    for nb in pim.get("nodeBlks", []) or []:
        nid = str(nb.get("Id"))
        x, y, _z = _xyz_from_str(nb.get("InsertionPoint", "0, 0, 0"))
        try:
            elev = float(nb.get("Height") or 0.0)
        except ValueError:
            elev = 0.0
        is_source = int(nb.get("IoNode", 0) or 0) == 1
        cn = CommonNode(
            id=nid,
            x=x,
            y=y,
            elevation_m=elev,
            kind="wt" if is_source else "base",
            raw={"has_label": nb.get("Label")},
        )
        if is_source:
            cn.pressure_bar = 1.013  # 개방형 수원 기본 (대기압); 실 압력은 펌프가 부여
        net.nodes[nid] = cn

    # ── nozzle → 참조 노드 kind 승격 + K
    for z in pim.get("nozzleBlks", []) or []:
        nid = str(z.get("SNodeId"))
        cn = net.nodes.get(nid)
        if cn is None:
            continue
        row = ktab.get(str(z.get("NozzleType")), {})
        k = row.get("KFactor")
        if k is None:
            # 테이블에 없으면 Min 유량/압력으로 역산: K = Q / √(P[MPa]×10)
            try:
                q = float(z.get("MinFlowQuantity") or 0)
                p = float(z.get("MinPressure") or 0)
                k = q / math.sqrt(p * 10.0) if p > 0 else None
            except (ValueError, ZeroDivisionError):
                k = None
        cn.kind = "nozzle"
        cn.k_factor_si = float(k) if k else None
        cn.raw["nozzle_type"] = z.get("NozzleType")

    # ── pump → 참조 노드 kind 승격 + 곡선
    for pmp in pim.get("pumpBlks", []) or []:
        nid = str(pmp.get("SNodeId"))
        cn = net.nodes.get(nid)
        if cn is None:
            continue
        try:
            maxq = float(pmp.get("MaxFlowQuantity") or 0)
        except ValueError:
            maxq = 0.0
        cn.kind = "pump"
        cn.pump_curve = {
            "rated_q": maxq / 1.5 if maxq else 0.0,
            "rated_p": 0.0,
            "shutoff_p": 0.0,
            "peak_q": maxq,
            "peak_p": 0.0,
            "curve_data": [],
        }
        cn.raw["pump_enode"] = pmp.get("ENodeId")
        cn.raw["pump_type"] = pmp.get("PumpType")

    # ── PRV(reducing) / orifice → kind 승격 (정보 보존; emit 에서 재구성)
    for rb in pim.get("reducingBlks", []) or []:
        cn = net.nodes.get(str(rb.get("SNodeId")))
        if cn is not None:
            cn.kind = "prv"
            cn.valve_type = "prv"
            cn.raw["prv_enode"] = rb.get("ENodeId")
            cn.raw["prv_block"] = rb
    for ob in pim.get("orifiaceBlks", []) or []:
        cn = net.nodes.get(str(ob.get("SNodeId")))
        if cn is not None:
            cn.kind = "orifice"
            cn.valve_type = "orifice"
            cn.raw["orifice_diameter"] = ob.get("Diameter")
            cn.raw["orifice_enode"] = ob.get("ENodeId")
            cn.raw["orifice_block"] = ob

    # ── 파이프
    for pl in pim.get("plines", []) or []:
        pid = str(pl.get("Id"))
        pts = [_xyz_from_str(s)[:2] for s in (pl.get("Points", []) or [])]
        waypoints = pts[1:-1] if len(pts) > 2 else []
        fittings: list[CommonFitting] = []
        for cnt_field, ftype in _CNT_TO_FITTING.items():
            try:
                c = int(float(pl.get(cnt_field, 0) or 0))
            except ValueError:
                c = 0
            if c > 0:
                fittings.append(CommonFitting(type_id=ftype, count=c))
        try:
            nominal = int(float(pl.get("PipeDiameter") or 0))
        except ValueError:
            nominal = 0
        try:
            length = float(pl.get("PipeLength") or 0.0)
        except ValueError:
            length = 0.0
        try:
            cfac = float(pl.get("CFactor") or 120)
        except ValueError:
            cfac = 120.0
        cp = CommonPipe(
            id=pid,
            start=str(pl.get("SNodeId")),
            end=str(pl.get("ENodeId")),
            length_m=length,
            diameter_inner_mm=0.0,
            nominal_mm=nominal,
            c_factor=cfac,
            pipe_type_label=str(pl.get("PipeName") or "KSD 3507"),
            fittings=fittings,
            waypoints=waypoints,
            raw={
                "has_label": pl.get("Label"),
                "material": pl.get("PipeMaterial"),
            },
        )
        net.pipes[pid] = cp

    proj = doc.get("Project", {}) or {}
    net.project_meta = {
        "title": proj.get("ProjectName", ""),
        "designer": proj.get("Designer", ""),
        "design_area": proj.get("Calculation", ""),
        "description": proj.get("Description", ""),
    }
    return net


# ────────────────────────────────────────────────────────────────────────────
# emit_has — CommonNetwork → .has (스켈레톤 기반)
# ────────────────────────────────────────────────────────────────────────────


def _zero_results(block: dict) -> dict:
    """블록 안 RESULT_* 키를 전부 "0" 으로 초기화 (HASS 가 재계산)."""
    for k in list(block.keys()):
        if str(k).startswith("RESULT_"):
            block[k] = "0"
    return block


# HASS 가 숫자로 파싱하는 문자열 필드. "" 를 받으면 파일 로드 실패(observed: nodeBlks.Height,
# pumpBlks.MaxFlowQuantity/MinFlowQuantity 가 "" 일 때 HASS 가 파일을 거부). 안전망으로 "0".
_HAS_NUMERIC_STRING_FIELDS = frozenset({
    "Height",
    "PipeLength", "PipeDiameter", "CFactor",
    "MaxFlowQuantity", "MinFlowQuantity",
    "MinPressure", "MaxPressure", "MinFlowRate",
    "Diameter", "KFactor", "Efficiency", "FixFlowRate",
    "Rotation",
})


def _sanitize_blank_numeric_strings(obj) -> None:
    """재귀: 알려진 숫자 필드의 "" → "0" in-place. PumpType/NozzleType 같은 문자열 필드는 손대지 않음.

    HASS 가 빈 문자열을 받으면 파일 로드 자체가 실패하므로, emit 마지막 단계에 한 번 돌려
    어떤 경로로 들어온 blank 든 막는다 (얇은 방어막).
    """
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, str) and v == "" and k in _HAS_NUMERIC_STRING_FIELDS:
                obj[k] = "0"
            elif isinstance(v, (dict, list)):
                _sanitize_blank_numeric_strings(v)
    elif isinstance(obj, list):
        for item in obj:
            _sanitize_blank_numeric_strings(item)


def _ensure_nozzle_type(db: dict, k: float) -> str:
    """nozzleDataTable 에서 KFactor 가 k 에 가장 가까운 NozzleName 반환.

    허용오차(±1) 밖이면 새 행을 추가하고 그 이름을 돌려준다 → parse 시 동일 K 복원.
    """
    table = db.setdefault("nozzleDataTable", [])
    best_name = None
    best_d = 1e9
    for row in table:
        kf = row.get("KFactor")
        if kf is None:
            continue
        d = abs(float(kf) - k)
        if d < best_d:
            best_d = d
            best_name = row.get("NozzleName")
    if best_name is not None and best_d <= 1.0:
        return str(best_name)
    # 새 행 추가 (round-trip 무손실)
    name = f"HEAD-K{k:g}"
    table.append({
        "Id": len(table) + 1,
        "NozzleName": name,
        "NozzleDec": name,
        "CalcMethod": "K-Factor",
        "KFactor": float(k),
        "FixFlowRate": None,
        "MinPressure": 0.1,
        "MaxPressure": 1.2,
        "MinFlowRate": round(float(k) * math.sqrt(1.0), 3),
    })
    return name


def _set_pump_flow_table(db: dict, pump_name: str, curve: dict) -> bool:
    """pumpFlowDataTable 의 ``pump_name`` 행을 성능곡선 3점으로 교체.

    화재안전기준 표준 곡선(체절 140% / 정격 / 150% 65%)을 HASS 가 H-Q 보간에 쓰는
    {PumpName, PumpFlowRate(L/min), PumpFlowHead(m)} 행으로 기록한다. curve 에 정격
    토출량/양정이 없으면(0) 스켈레톤 기본 곡선을 그대로 두고 False 를 돌려준다.

    curve 키(parse_sdf Pump-fan attribute 유래) — 압력은 **bar 로 정규화**되어 옴
    (parse_sdf 가 rated-p-unit 으로 통일). HASS PumpFlowHead 는 양정[m] 이므로
    bar→m(×_BAR_TO_M) 변환해 기록한다:
        rated_q  = 정격 토출량 (L/min)
        rated_p  = 정격 압력 (bar) → 양정[m]
        shutoff_p= 체절 압력 (bar) → 양정[m]
        peak_q   = 150% 토출량 (L/min)
        peak_p   = 150% 압력 (bar) → 양정[m]
    """
    q = float(curve.get("rated_q") or 0.0)
    h = float(curve.get("rated_p") or 0.0) * _BAR_TO_M
    if q <= 0.0 or h <= 0.0:
        return False
    shut = (float(curve.get("shutoff_p") or 0.0) * _BAR_TO_M) or round(h * 1.40, 3)
    pq = float(curve.get("peak_q") or 0.0) or round(q * 1.50, 3)
    ph = (float(curve.get("peak_p") or 0.0) * _BAR_TO_M) or round(h * 0.65, 3)
    points = [(0.0, shut), (q, h), (pq, ph)]

    table = db.setdefault("pumpFlowDataTable", [])
    # 같은 PumpName 의 기존 행 제거 후 재구성 (스켈레톤 FP 기본곡선 대체)
    table[:] = [r for r in table if str(r.get("PumpName")) != str(pump_name)]
    for fr, fh in points:
        table.append({
            "Id": len(table) + 1,
            "PumpName": str(pump_name),
            "PumpFlowRate": round(float(fr), 3),
            "PumpFlowHead": round(float(fh), 3),
        })
    return True


def emit_has(
    net: CommonNetwork,
    path: str | Path | None = None,
    *,
    isometric: bool = False,
    iso_z_scale: float = 1.0,
) -> dict:
    """CommonNetwork → HAS dict(JSON). path 주어지면 파일에도 UTF-8 로 기록.

    스켈레톤(붙임 샘플)을 로드해 참조 DB/범례/설정/단위계를 그대로 쓰고,
    ``PipeInfoMgr`` 의 도면+해석 블록만 우리 네트워크로 교체한다. 원본 도면 잔재
    (infoTxts/lines/장식 심볼 등)는 전부 비워 source 프로젝트 정보 누출을 막는다.

    ``isometric=True`` 면 노드/배관의 화면좌표(InsertionPoint/Points X,Y)를
    평면(x,y)+고도(elevation)의 **30° 등각투영**으로 베이크해 HASS/solver 화면처럼
    계통도(아이소뷰)로 보이게 한다. 평면뷰는 층이 겹쳐 라이저가 점으로 뭉치지만,
    등각투영은 고도를 화면 수직축으로 펼쳐 층·입상관이 분리돼 보인다. **순수 표시
    변환** — Height/PipeLength/흐름방향은 그대로라 수리계산 결과엔 영향 0.
    ``iso_z_scale`` 은 고도 펼침 배율(1.0=고도범위가 평면 대각선의 절반).
    """
    skel = json.loads(_template_path().read_text(encoding="utf-8-sig"))
    doc = skel["Document"]
    pim = doc["PipeInfoMgr"]
    db = doc.setdefault("PipeDbMgr", {})

    # 필드 완전성용 prototype (HASS 가 쓴 모든 키 보존) — 데이터만 덮어씀
    proto_node = copy.deepcopy(pim["nodeBlks"][0])
    proto_pline = copy.deepcopy(pim["plines"][0])
    proto_nozzle = copy.deepcopy(pim["nozzleBlks"][0]) if pim.get("nozzleBlks") else {}
    proto_pump = copy.deepcopy(pim["pumpBlks"][0]) if pim.get("pumpBlks") else {}
    proto_reducing = copy.deepcopy(pim["reducingBlks"][0]) if pim.get("reducingBlks") else {}
    proto_orifice = copy.deepcopy(pim["orifiaceBlks"][0]) if pim.get("orifiaceBlks") else {}
    # ★ directionBlk(흐름방향 화살표)은 원본 정보가 아니라 pline 마다 1:1 로 필요한
    #   동반 엔티티. HASS 뷰어가 파일을 열 때 각 pline 의 방향블록을 참조하므로
    #   비워두면 "개체 참조가 설정되지 않았습니다" null ref 로 열리지 않는다.
    #   (붙임 레퍼런스 4종 모두 plines:directionBlks = 1:1, ParentId⊆pline Id)
    proto_direction = copy.deepcopy(pim["directionBlks"][0]) if pim.get("directionBlks") else {}

    nodes = list(net.nodes.values())
    base_scale = _coord_scale(nodes)
    # schematic 압축 보정 — 뭉친 망을 레퍼런스 수준 간격으로 균일 확대(시각화 전용).
    scale = base_scale * _spread_factor(nodes, base_scale)

    def _dz(cn) -> float:
        # 표시 전용 z — 라이저 입상관 높이·헤드평면 평탄화(display_z_m). 없으면 실표고.
        # HAS 등각 lift 는 이 표시 z 로(=계통도 형상), Height(수리표고)는 elevation_m 유지.
        v = cn.raw.get("display_z_m")
        return cn.elevation_m if v is None else v

    # ── 등각투영(계통도) 화면좌표 변환 ─────────────────────────────────────────
    # 평면(px,py)+고도(elev) → 30° 등각투영 화면좌표. 고도 펼침은 평면 표시 대각선에
    # 정규화해 좌표 단위(m/mm)·_spread_factor 배율과 무관하게 비례를 유지한다.
    # 순수 표시 변환(Height/PipeLength/방향 불변) — 수리계산 영향 0.
    _COS30, _SIN30 = 0.8660254037844387, 0.5
    _elevs = [_dz(c) for c in nodes if _dz(c) is not None]
    _e_mid = (min(_elevs) + max(_elevs)) / 2.0 if _elevs else 0.0
    _e_range = (max(_elevs) - min(_elevs)) if _elevs else 0.0
    if isometric and _elevs:
        _pxs = [c.x * scale for c in nodes]
        _pys = [c.y * scale for c in nodes]
        _diag = math.hypot(
            (max(_pxs) - min(_pxs)) if _pxs else 0.0,
            (max(_pys) - min(_pys)) if _pys else 0.0,
        )
        _lift_per_m = (_diag * 0.5 * iso_z_scale / _e_range) if _e_range > 0 else 0.0
    else:
        _lift_per_m = 0.0

    def _proj(px: float, py: float, elev: float | None) -> tuple[float, float]:
        """평면 표시좌표(px,py)+고도 → 등각투영 화면(X,Y). isometric off 면 그대로."""
        if not isometric:
            return px, py
        e = _e_mid if elev is None else elev
        return (
            (px - py) * _COS30,
            (px + py) * _SIN30 + (e - _e_mid) * _lift_per_m,
        )

    # 단일 source 선택: wt 우선 → pump → 첫 노드
    source_id: str | None = None
    for cn in nodes:
        if cn.kind == "wt":
            source_id = cn.id
            break
    if source_id is None:
        for cn in nodes:
            if cn.kind == "pump":
                source_id = cn.id
                break
    if source_id is None and nodes:
        source_id = nodes[0].id

    # ── 흐름방향 정규화 ──────────────────────────────────────────────────────
    # HASS 입력검사는 배관이 source(수원)에서 멀어지는 방향(SNodeId=상류, ENodeId=하류)
    # 으로 설치돼야 통과한다. 우리 망의 start/end 는 DXF 끝점/그래프 순서라 임의 방향이
    # 라서, source 에서 멀어지는 절반가량 파이프가 "배관 역방향 설치" 입력값 오류로 막힌다.
    # source 로부터 BFS 깊이를 구해 start 가 더 깊으면(=하류면) start↔end 를 뒤집는다.
    # 길이·직경·CFactor 불변이라 수리계산 영향 0 — 입력 방향 메타와 흐름화살표만 교정.
    # 이 교정은 pline 루프보다 먼저 실행해야 SNodeId/ENodeId 와 directionBlk(start→end
    # 각도)가 함께 정합된다.
    if source_id is not None and net.pipes:
        from collections import deque as _deque
        _adj: dict[str, list[str]] = {}
        for _cp in net.pipes.values():
            _adj.setdefault(_cp.start, []).append(_cp.end)
            _adj.setdefault(_cp.end, []).append(_cp.start)
        # 2-port 인라인 요소(펌프·PRV·오리피스)는 net.pipes 에 없는 별도 엣지로
        # 망에 끼어든다. 펌프 흡입측 노드(=Input/수원, 흔히 source_id)는 토출 노드와
        # *오직 펌프 요소로만* 이어져 파이프 인접에 안 잡힌다. 이 엣지를 BFS 인접에
        # 넣지 않으면 source 에서 망 본체로 못 건너가 depth 가 전부 None → 한 건도
        # 안 뒤집혀 "역방향 설치" 오류가 그대로 남는다(펌프 추가 후 재현된 버그).
        for _cn in nodes:
            _out = None
            if _cn.kind == "pump":
                _out = _cn.raw.get("pump_output")
            elif _cn.kind == "prv":
                _out = _cn.raw.get("prv_enode")
            elif _cn.kind == "orifice":
                _out = _cn.raw.get("orifice_enode")
            if _out is not None:
                _out = str(_out)
                _adj.setdefault(_cn.id, []).append(_out)
                _adj.setdefault(_out, []).append(_cn.id)
        _depth: dict[str, int] = {source_id: 0}
        _dq = _deque([source_id])
        while _dq:
            _u = _dq.popleft()
            for _v in _adj.get(_u, ()):
                if _v not in _depth:
                    _depth[_v] = _depth[_u] + 1
                    _dq.append(_v)
        _flipped = 0
        for _cp in net.pipes.values():
            _ds = _depth.get(_cp.start)
            _de = _depth.get(_cp.end)
            if _ds is not None and _de is not None and _ds > _de:
                _cp.start, _cp.end = _cp.end, _cp.start
                if _cp.waypoints:
                    _cp.waypoints = list(reversed(_cp.waypoints))
                _flipped += 1

    # 전역 Id 카운터 (노드·pline·nozzle·pump 가 한 Id 공간 공유)
    gid = 1
    idmap: dict[str, int] = {}
    xy_m: dict[str, tuple[float, float]] = {}

    node_blks: list[dict] = []
    for cn in nodes:
        hid = gid
        gid += 1
        idmap[cn.id] = hid
        x, y = _proj(cn.x * scale, cn.y * scale, _dz(cn))
        xy_m[cn.id] = (x, y)
        nb = copy.deepcopy(proto_node)
        nb["Id"] = hid
        nb["Label"] = str(cn.raw.get("has_label") or cn.id)
        nb["InsertionPoint"] = _xyz_to_str(x, y, 0.0)
        nb["Height"] = f"{cn.elevation_m:g}" if cn.elevation_m is not None else "0"
        nb["IoNode"] = 1 if cn.id == source_id else 0
        node_blks.append(_zero_results(nb))

    # plines
    pline_blks: list[dict] = []
    pline_geom: list[tuple[int, float, float, float, float]] = []  # (pid, sx,sy, ex,ey)
    for cp in net.pipes.values():
        pid = gid
        gid += 1
        pl = copy.deepcopy(proto_pline)
        pl["Id"] = pid
        pl["Label"] = str(cp.raw.get("has_label") or cp.id)
        pl["SNodeId"] = str(idmap.get(cp.start, cp.start))
        pl["ENodeId"] = str(idmap.get(cp.end, cp.end))
        # Points: start → waypoints → end (m)
        sx, sy = xy_m.get(cp.start, (0.0, 0.0))
        ex, ey = xy_m.get(cp.end, (0.0, 0.0))
        pline_geom.append((pid, sx, sy, ex, ey))
        # 경유점(bend)은 고도값이 없어 양 끝 노드의 중간 고도로 투영 — 등각에서 직선 유지.
        _se = net.nodes.get(cp.start)
        _ee = net.nodes.get(cp.end)
        _wp_elev = None
        if isometric:
            _ev = [_dz(n) for n in (_se, _ee)
                   if n is not None and _dz(n) is not None]
            _wp_elev = (sum(_ev) / len(_ev)) if _ev else _e_mid
        pts = [_xyz_to_str(sx, sy)]
        for wx, wy in cp.waypoints:
            pts.append(_xyz_to_str(*_proj(wx * scale, wy * scale, _wp_elev)))
        pts.append(_xyz_to_str(ex, ey))
        pl["Points"] = pts
        pl["PipeDiameter"] = str(cp.nominal_mm or int(round(cp.diameter_inner_mm)))
        pl["PipeLength"] = f"{cp.length_m:.3f}"
        pl["CFactor"] = str(int(round(cp.c_factor)))
        # fitting 카운트 — 전부 0 으로 리셋 후 매핑된 것만 채움
        for f in _CNT_FIELDS:
            pl[f] = "0"
        for fit in cp.fittings:
            cnt_field = _FITTING_TO_CNT.get(str(fit.type_id).lower())
            if cnt_field:
                pl[cnt_field] = str(int(pl[cnt_field]) + int(fit.count or 1))
        pline_blks.append(_zero_results(pl))

    # nozzleBlks — kind=nozzle 노드마다
    nozzle_blks: list[dict] = []
    if proto_nozzle:
        for cn in nodes:
            if cn.kind not in ("nozzle", "head"):
                continue
            nid = gid
            gid += 1
            nz = copy.deepcopy(proto_nozzle)
            nz["Id"] = nid
            nz["Label"] = str(cn.raw.get("has_label") or f"H{idmap[cn.id]}")
            nz["SNodeId"] = str(idmap[cn.id])
            nx, ny = xy_m.get(cn.id, (0.0, 0.0))
            nz["InsertionPoint"] = _xyz_to_str(nx, ny, 0.0)
            nz["FlwoCalMethod"] = "K-Factor"
            if cn.k_factor_si:
                nz["NozzleType"] = _ensure_nozzle_type(db, float(cn.k_factor_si))
            nozzle_blks.append(_zero_results(nz))

    # pumpBlks — kind=pump 노드마다
    pump_blks: list[dict] = []
    if proto_pump:
        for cn in nodes:
            if cn.kind != "pump":
                continue
            pid = gid
            gid += 1
            pb = copy.deepcopy(proto_pump)
            pb["Id"] = pid
            pb["Label"] = str(cn.raw.get("has_label") or "FP")
            pb["SNodeId"] = str(idmap[cn.id])
            # ENodeId = 펌프 토출 노드. 펌프는 SRC→토출의 유일한 연결(병렬 파이프 없음)
            # 이라 파이프로 추론 불가 → parse_sdf 가 보존한 pump_output 을 우선 사용.
            # (없으면 파이프 이웃 추론, 그조차 없으면 self — 구버전 호환)
            pump_out = cn.raw.get("pump_output")
            enode = idmap.get(str(pump_out)) if pump_out is not None else None
            if enode is None:
                enode = _downstream_hid(net, idmap, cn.id)
            pb["ENodeId"] = str(enode if enode is not None else idmap[cn.id])
            px, py = xy_m.get(cn.id, (0.0, 0.0))
            pb["InsertionPoint"] = _xyz_to_str(px, py, 0.0)
            curve = cn.pump_curve or {}
            peak_q = curve.get("peak_q") or curve.get("rated_q") or 0
            # HASS 가 거부하는 "" 을 막기 위해 PumpType 명시 set (proto_pump 의 값 보존, 없으면 "FP").
            # MaxFlowQuantity 는 0 이라도 "0" 문자열은 들어가도록 강제 (peak_q 가 0 일 때 str(0)="0").
            pb["PumpType"] = (pb.get("PumpType") or cn.raw.get("pump_type") or "FP")
            pb["MaxFlowQuantity"] = str(peak_q) if str(peak_q).strip() else "0"
            pb["MinFlowQuantity"] = "0"
            # 사용자 성능곡선이 있으면 pumpFlowDataTable 의 해당 PumpType 행을 교체
            # (없으면 스켈레톤 기본 FP 곡선 유지). PumpType ↔ PumpName 일치 필수.
            _set_pump_flow_table(db, pb["PumpType"], curve)
            pump_blks.append(_zero_results(pb))

    # reducingBlks(PRV) — kind=prv 노드마다 (2-port 인라인 감압변)
    reducing_blks: list[dict] = []
    if proto_reducing:
        for cn in nodes:
            if cn.kind != "prv":
                continue
            rid = gid
            gid += 1
            rb = copy.deepcopy(cn.raw.get("prv_block") or proto_reducing)
            rb["Id"] = rid
            rb["SNodeId"] = str(idmap[cn.id])
            en = cn.raw.get("prv_enode")
            rb["ENodeId"] = str(idmap.get(str(en), en) if en is not None
                                else _downstream_hid(net, idmap, cn.id) or idmap[cn.id])
            rx, ry = xy_m.get(cn.id, (0.0, 0.0))
            rb["InsertionPoint"] = _xyz_to_str(rx, ry, 0.0)
            reducing_blks.append(_zero_results(rb))

    # orifiaceBlks(오리피스) — kind=orifice 노드마다
    orifice_blks: list[dict] = []
    if proto_orifice:
        for cn in nodes:
            if cn.kind != "orifice":
                continue
            oid = gid
            gid += 1
            ob = copy.deepcopy(cn.raw.get("orifice_block") or proto_orifice)
            ob["Id"] = oid
            ob["SNodeId"] = str(idmap[cn.id])
            en = cn.raw.get("orifice_enode")
            ob["ENodeId"] = str(idmap.get(str(en), en) if en is not None
                                else _downstream_hid(net, idmap, cn.id) or idmap[cn.id])
            ox, oy = xy_m.get(cn.id, (0.0, 0.0))
            ob["InsertionPoint"] = _xyz_to_str(ox, oy, 0.0)
            orifice_blks.append(_zero_results(ob))

    # directionBlks — pline 마다 1:1 흐름방향 화살표 (ParentId=pline Id).
    # InsertionPoint=pline 중점, Rotation=start→end 방향각. HASS 가 열 때 필수.
    direction_blks: list[dict] = []
    if proto_direction:
        for pid, sx, sy, ex, ey in pline_geom:
            did = gid
            gid += 1
            d = copy.deepcopy(proto_direction)
            d["Id"] = did
            d["ParentId"] = pid
            d["InsertionPoint"] = _xyz_to_str((sx + ex) / 2.0, (sy + ey) / 2.0, 0.0)
            d["Rotation"] = math.atan2(ey - sy, ex - sx)
            d["IsPositive"] = True
            direction_blks.append(d)

    # PipeInfoMgr 교체 — geometry 만 우리 것, 나머지(원본 도면 잔재)는 비움
    pim["nodeBlks"] = node_blks
    pim["plines"] = pline_blks
    pim["nozzleBlks"] = nozzle_blks
    pim["pumpBlks"] = pump_blks
    pim["reducingBlks"] = reducing_blks
    pim["orifiaceBlks"] = orifice_blks
    pim["directionBlks"] = direction_blks
    for key in (
        "lines", "angleBlks", "alarmBlks", "butterBlks",
        "delugeBlks", "dryBlks", "flexibleBlks", "gateBlks", "lengthPartBlks",
        "swingcheckBlks", "strainerBlks", "preactionBlks",
        "flowFactorBlks", "pressureLossBlks", "infoTxts", "txts",
    ):
        if key in pim:
            pim[key] = []

    # Project 메타 덮어쓰기
    proj = doc.setdefault("Project", {})
    meta = net.project_meta or {}
    if meta.get("title"):
        proj["ProjectName"] = meta["title"]
    elif not proj.get("ProjectName"):
        proj["ProjectName"] = "Remote30 통합 배관망"
    if meta.get("designer"):
        proj["Designer"] = meta["designer"]
    if meta.get("design_area"):
        proj["Calculation"] = meta["design_area"]

    out = {"Document": doc}
    # 최종 안전망 — 어떤 경로로든 들어온 빈 숫자 문자열을 "0" 으로 채움.
    # HASS 는 nodeBlks.Height / pumpBlks.MaxFlowQuantity 등 숫자 파싱 필드의
    # "" 을 받으면 파일 로드 자체가 실패한다 (관측: 일예시 파.has 8건 diff).
    _sanitize_blank_numeric_strings(out)
    if path is not None:
        Path(path).write_text(
            json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    return out


def _downstream_hid(net: CommonNetwork, idmap: dict[str, int], node_id: str) -> int | None:
    """node_id 에 연결된 파이프의 반대편 노드 HAS Id (펌프 ENodeId 용)."""
    for cp in net.pipes.values():
        if cp.start == node_id and cp.end in idmap:
            return idmap[cp.end]
        if cp.end == node_id and cp.start in idmap:
            return idmap[cp.start]
    return None


# ────────────────────────────────────────────────────────────────────────────
# SDF ↔ HAS 변환 + round-trip
# ────────────────────────────────────────────────────────────────────────────


def convert_sdf_to_has(
    sdf_path: str | Path,
    has_path: str | Path,
    *,
    isometric: bool = False,
    iso_z_scale: float = 1.0,
    simplify: bool = True,
) -> Path:
    """SDF → HAS. parse_sdf 로 CommonNetwork 만든 뒤 emit_has.

    ``isometric=True`` 면 화면좌표를 등각투영(계통도)으로 베이크 — emit_has 참고.
    ``simplify`` — 일직선 통과 base 노드 병합으로 출력 노드 수 절감(무손실). 기본 True.
    """
    net = parse_sdf(sdf_path)
    if simplify:
        simplify_passthrough_nodes(net)
    emit_has(net, has_path, isometric=isometric, iso_z_scale=iso_z_scale)
    return Path(has_path)


def convert_has_to_sdf(has_path: str | Path, sdf_path: str | Path) -> Path:
    """HAS → SDF. parse_has → emit_sdf_xml."""
    net = parse_has(has_path)
    xml = emit_sdf_xml(net)
    Path(sdf_path).write_text(xml, encoding="utf-8")
    return Path(sdf_path)


def parse_for_preview(path: str | Path) -> dict:
    """캔버스 미리보기용 경량 표현 — {nodes, edges, bbox, meta, format}."""
    net = parse_has(path)
    nodes = [
        {"id": n.id, "x": n.x, "y": n.y, "kind": n.kind,
         "label": n.raw.get("has_label", n.id)}
        for n in net.nodes.values()
    ]
    edges = [{"start": p.start, "end": p.end, "nominal_mm": p.nominal_mm}
             for p in net.pipes.values()]
    xs = [n["x"] for n in nodes] or [0.0]
    ys = [n["y"] for n in nodes] or [0.0]
    return {
        "nodes": nodes,
        "edges": edges,
        "bbox": {"min_x": min(xs), "min_y": min(ys),
                 "max_x": max(xs), "max_y": max(ys)},
        "meta": net.project_meta,
        "format": "has",
    }


def round_trip_check(has_path: str | Path) -> dict:
    """HAS → CommonNetwork → HAS 무손실 검증 — 노드/파이프 수·kind·길이 비교."""
    net1 = parse_has(has_path)
    out = emit_has(net1)
    net2 = has_dict_to_network(out)

    def _kind_dist(n: CommonNetwork) -> dict[str, int]:
        d: dict[str, int] = {}
        for cn in n.nodes.values():
            d[cn.kind] = d.get(cn.kind, 0) + 1
        return d

    len1 = {p.id: p.length_m for p in net1.pipes.values()}
    # net2 는 새 Id 라 길이 분포만 비교
    l1 = sorted(len1.values())
    l2 = sorted(p.length_m for p in net2.pipes.values())
    if l1 and len(l1) == len(l2):
        rmse = math.sqrt(sum((a - b) ** 2 for a, b in zip(l1, l2)) / len(l1))
    else:
        rmse = float("nan")
    return {
        "nodes_in": len(net1.nodes),
        "nodes_out": len(net2.nodes),
        "pipes_in": len(net1.pipes),
        "pipes_out": len(net2.pipes),
        "kind_dist_in": _kind_dist(net1),
        "kind_dist_out": _kind_dist(net2),
        "length_rmse_m": round(rmse, 6) if rmse == rmse else None,
        "lossless": (
            len(net1.nodes) == len(net2.nodes)
            and len(net1.pipes) == len(net2.pipes)
            and _kind_dist(net1) == _kind_dist(net2)
        ),
    }


if __name__ == "__main__":
    import sys

    if len(sys.argv) >= 2:
        rep = round_trip_check(sys.argv[1])
        print(json.dumps(rep, ensure_ascii=False, indent=2))
    else:
        print("usage: python has_converter.py <file.has>   # round-trip 검증")

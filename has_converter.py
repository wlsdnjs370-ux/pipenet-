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
    emit_sdf_xml,
    parse_sdf,
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


def _coord_scale(nodes: list[CommonNode]) -> float:
    """좌표를 미터로 정규화하는 배율. bbox span 이 수 km 면 mm 로 보고 0.001.

    HAS 도면은 m 단위(예: -25.98, 11.25). 통합망(SDF)이 schematic mm 좌표면
    그대로 쓰면 도면이 1000배 커져 보임(수리계산엔 무영향 — PipeLength 사용).
    """
    xs = [n.x for n in nodes]
    ys = [n.y for n in nodes]
    if not xs:
        return 1.0
    span = max(max(xs) - min(xs), max(ys) - min(ys))
    return 0.001 if span > 5000.0 else 1.0


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


def emit_has(net: CommonNetwork, path: str | Path | None = None) -> dict:
    """CommonNetwork → HAS dict(JSON). path 주어지면 파일에도 UTF-8 로 기록.

    스켈레톤(붙임 샘플)을 로드해 참조 DB/범례/설정/단위계를 그대로 쓰고,
    ``PipeInfoMgr`` 의 도면+해석 블록만 우리 네트워크로 교체한다. 원본 도면 잔재
    (infoTxts/lines/장식 심볼 등)는 전부 비워 source 프로젝트 정보 누출을 막는다.
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

    nodes = list(net.nodes.values())
    scale = _coord_scale(nodes)

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

    # 전역 Id 카운터 (노드·pline·nozzle·pump 가 한 Id 공간 공유)
    gid = 1
    idmap: dict[str, int] = {}
    xy_m: dict[str, tuple[float, float]] = {}

    node_blks: list[dict] = []
    for cn in nodes:
        hid = gid
        gid += 1
        idmap[cn.id] = hid
        x, y = cn.x * scale, cn.y * scale
        xy_m[cn.id] = (x, y)
        nb = copy.deepcopy(proto_node)
        nb["Id"] = hid
        nb["Label"] = str(cn.raw.get("has_label") or cn.id)
        nb["InsertionPoint"] = _xyz_to_str(x, y, 0.0)
        nb["Height"] = f"{cn.elevation_m:g}"
        nb["IoNode"] = 1 if cn.id == source_id else 0
        node_blks.append(_zero_results(nb))

    # plines
    pline_blks: list[dict] = []
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
        pts = [_xyz_to_str(sx, sy)]
        for wx, wy in cp.waypoints:
            pts.append(_xyz_to_str(wx * scale, wy * scale))
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
            # ENodeId = pump 노드의 하류(또는 상류) 이웃
            enode = _downstream_hid(net, idmap, cn.id)
            pb["ENodeId"] = str(enode if enode is not None else idmap[cn.id])
            px, py = xy_m.get(cn.id, (0.0, 0.0))
            pb["InsertionPoint"] = _xyz_to_str(px, py, 0.0)
            curve = cn.pump_curve or {}
            peak_q = curve.get("peak_q") or curve.get("rated_q") or 0
            pb["MaxFlowQuantity"] = str(peak_q)
            pb["MinFlowQuantity"] = "0"
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

    # PipeInfoMgr 교체 — geometry 만 우리 것, 나머지(원본 도면 잔재)는 비움
    pim["nodeBlks"] = node_blks
    pim["plines"] = pline_blks
    pim["nozzleBlks"] = nozzle_blks
    pim["pumpBlks"] = pump_blks
    pim["reducingBlks"] = reducing_blks
    pim["orifiaceBlks"] = orifice_blks
    for key in (
        "lines", "directionBlks", "angleBlks", "alarmBlks", "butterBlks",
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


def convert_sdf_to_has(sdf_path: str | Path, has_path: str | Path) -> Path:
    """SDF → HAS. parse_sdf 로 CommonNetwork 만든 뒤 emit_has."""
    net = parse_sdf(sdf_path)
    emit_has(net, has_path)
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

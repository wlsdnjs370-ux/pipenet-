"""KFP ↔ SDF 양방향 변환기.

KFP (K-solver) 와 SDF (PIPENET Vision) 는 둘 다 한국 소방 SP 시스템의 hydraulic
모델을 표현하지만 schema 구조가 다르다:

* KFP — JSON, flat 구조. 노드/파이프/라이브러리를 top-level dict.
  핵심 키: nodes (좌표), nodes_meta (속성), pipe_data, nozzle_library, fittings_library,
  pump_library, design_settings, version "2.2-NODE-SCHEMA".
* SDF — XML, 계층 구조. Project > Network-spray > {Nodes, Links, Specifications}.
  Node 는 io-node 로 boundary 표시. Nozzle/Pipe-set/Fitting 별도 element.

이 모듈은 의미론적 동등성 (밸브/배관/헤드/피팅의 종류와 연결관계) 을 보존하면서
양방향 변환을 수행. 완전 round-trip 무손실은 어려움 — SDF 의 PIPENET 전용 메타
(Pipe-set 스키마, Calculation-spec) 와 KFP 의 hydraulic 결과 캐시 (flow_lpm 등) 는
도메인 외 메타라 단방향 손실 가능.

매핑 표 (의미 보존 코어):

| 도메인 항목       | KFP                                       | SDF                                                |
|------------------|-------------------------------------------|----------------------------------------------------|
| 노드 좌표         | nodes[id] = [x, y, z]                     | Node><Position x= y=/>                             |
| 노드 ID/라벨      | "N5" 등                                   | label="5"                                          |
| 노드 종류         | nodes_meta[id].type_id                    | io-node 와 별도 element (Nozzle/Pump-fan 등)        |
|  - base          | "base"                                    | io-node="No"                                       |
|  - nozzle/head   | "nozzle" + k_factor_si                    | <Nozzle input=...> + nodes_meta 의 k 별도 spec     |
|  - wt (수원)     | "wt"                                      | io-node="Input" + <Calculation-spec pressure=/>    |
|  - valve         | "valve"                                   | <Elastomeric-valve input=... output=...>           |
|  - pump          | "pump" + pump_library_id                  | <Pump-fan input= output= library_pump=>            |
| 노드 elev        | nodes_meta[id].elevation_m                | elevation 속성                                     |
| 파이프 연결       | pipe_data[id].{start, end}                | Pipe input= output=                                |
| 파이프 직경 (내)  | pipe_data[id].diameter                    | Pipe bore (m)                                      |
| 파이프 호칭경     | pipe_data[id].nominal_mm                  | Pipe-type 안 Pipe-size 와 매핑                     |
| 파이프 길이       | pipe_data[id].length_m                    | Pipe length                                        |
| C-factor         | pipe_data[id].C                           | Pipe-type c-factor                                 |
| 등가길이 추가     | pipe_data[id].equivalent_length           | Pipe 안 Fitting count + type                       |
| Fitting 종류      | fittings 리스트 (L/D)                     | Fitting type="elbow|elbow-45|tee" count=N         |

엔트리 포인트:
    convert_kfp_to_sdf(kfp_path) → sdf_xml_str
    convert_sdf_to_kfp(sdf_path) → kfp_dict (JSON dumpable)
    parse_for_preview(path) → {nodes:[{id,x,y,z,kind}], edges:[{id,a,b,dn,length}], meta:{}}

사용 예:
    sdf_xml = convert_kfp_to_sdf("path/to/file.kfp")
    Path("out.sdf").write_text(sdf_xml, encoding="utf-8")
"""

from __future__ import annotations

import copy
import datetime
import json
import math
import os
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path


# 1 bar(=100000 Pa) ↔ 물기둥 양정[m]. ρg=9806.65 (remote30_full_network.M_TO_PA 와 동일).
#   CommonNetwork.pump_curve 의 압력(rated_p/shutoff_p/peak_p)은 내부적으로 bar 로
#   정규화한다(KFP 네이티브 단위). SDF Pump-fan rated-p 는 생산자마다 단위가 달라
#   (KFP→SDF=bar, 통합망=양정[m]) rated-p-unit 속성으로 구분해 parse_sdf 가 bar 로 통일.
_BAR_TO_M = 100000.0 / 9806.65   # ≈ 10.19716  bar → m
_M_TO_BAR = 9806.65 / 100000.0   # ≈ 0.0980665 m → bar


# ────────────────────────────────────────────────────────────────────────────
# 도메인 모델 — KFP/SDF 공통의 의미론적 표현
# ────────────────────────────────────────────────────────────────────────────


@dataclass
class CommonNode:
    """도메인 중립 노드 — KFP/SDF 양쪽이 매핑되는 공통 표현."""

    id: str                          # KFP 의 "N5", SDF 의 label="5" (양쪽 호환 위해 KFP 원형 보존)
    x: float                         # m 단위
    y: float
    elevation_m: float = 0.0         # z 좌표 / SDF elevation
    kind: str = "base"               # "base" | "nozzle" | "wt" | "valve" | "pump"
    k_factor_si: float | None = None # nozzle 인 경우 K (SI = L/min·bar^-0.5)
    pressure_bar: float | None = None  # wt (수원) 의 boundary 압력
    pump_curve: dict | None = None   # pump 인 경우 곡선 데이터
    valve_type: str | None = None    # valve 인 경우 종류 (alarm/check/gate/...)
    is_check_valve: bool = False
    raw: dict = field(default_factory=dict)   # 변환 손실 방지용 원본 보관


@dataclass
class CommonFitting:
    """파이프 안의 피팅 (elbow, tee 등)."""

    type_id: str        # "elbow" | "elbow-45" | "tee" | "gate_valve" | ...
    count: int = 1
    l_over_d: float | None = None  # KFP 의 L/D 값 (등가길이 비). SDF 는 type 이름만 저장.


@dataclass
class CommonPipe:
    """도메인 중립 파이프."""

    id: str                          # KFP "P7", SDF Pipe label
    start: str                       # 노드 id (start side)
    end: str                         # 노드 id (end side)
    length_m: float = 0.0
    diameter_inner_mm: float = 0.0   # 실제 내경 (KFP diameter)
    nominal_mm: int = 0              # 호칭경 (KFP nominal_mm)
    c_factor: float = 120.0
    roughness_mm: float = 0.15
    pipe_type_label: str = "KSD 3507"  # KFP type 또는 SDF Pipe-type Name
    fittings: list[CommonFitting] = field(default_factory=list)
    equivalent_length_m: float = 0.0
    # PIPENET SDF 의 <Waypoints> — 본관이 ㄷ자로 휘어진 경우 중간 꺾임점들.
    # start → wp1 → wp2 → ... → end 폴리라인 경로 (m 단위).
    # 미리보기에서 직선 대신 폴리라인으로 그려야 도면 모양 정확.
    waypoints: list[tuple[float, float]] = field(default_factory=list)
    raw: dict = field(default_factory=dict)


@dataclass
class CommonNetwork:
    """KFP/SDF 공통 hydraulic 네트워크 표현."""

    nodes: dict[str, CommonNode] = field(default_factory=dict)
    pipes: dict[str, CommonPipe] = field(default_factory=dict)
    nozzle_library: list[dict] = field(default_factory=list)   # K-factor 카탈로그
    fitting_library: list[dict] = field(default_factory=list)  # L/D 카탈로그
    pump_library: list[dict] = field(default_factory=list)     # 펌프 곡선 카탈로그
    design_settings: dict = field(default_factory=dict)        # min_p, calc_method, etc.
    project_meta: dict = field(default_factory=dict)           # title, company, ...
    source_format: str = ""                                    # "kfp" | "sdf"


# ────────────────────────────────────────────────────────────────────────────
# 노드 수 최적화 — 일직선 통과 노드(pass-through) 병합
#   입상관(라이저)·직선 본관은 중간 노드가 여러 개일 필요가 없다(첫·끝 2개면 족함).
#   "수리 무손실"인 경우에만 병합한다:
#     · kind == "base" (헤드/노즐/수원/펌프/밸브/PRV/오리피스는 절대 제거 안 함)
#     · degree == 2 (정확히 두 배관만 접속)
#     · 두 배관의 호칭경·재질(type)·C 값 동일
#     · 세 점(a–n–b)이 3D 로 일직선 (라이저=수직, 직선본관=수평)
#   일직선이라 병합 후 직선거리 == 두 구간 길이합 → K-Fire 의
#   "좌표거리 == length_m" 불변식 보존, 도면 형상도 그대로(꺾임 없음).
#   꺾이는 본관(ㄷ자)은 일직선이 아니므로 보존된다(KFP 는 waypoint 직렬화가 없어
#   코너 노드를 지우면 형상·길이 둘 다 깨지기 때문).
# ────────────────────────────────────────────────────────────────────────────


def _collinear_3d(na: "CommonNode", nn: "CommonNode", nb: "CommonNode",
                  tol_deg: float) -> bool:
    """a–n–b 가 n 을 통과해 일직선인가 (∠a-n-b ≈ 180°). z 는 elevation_m 사용."""
    v1 = (na.x - nn.x, na.y - nn.y, na.elevation_m - nn.elevation_m)
    v2 = (nb.x - nn.x, nb.y - nn.y, nb.elevation_m - nn.elevation_m)
    m1 = math.sqrt(v1[0] ** 2 + v1[1] ** 2 + v1[2] ** 2)
    m2 = math.sqrt(v2[0] ** 2 + v2[1] ** 2 + v2[2] ** 2)
    if m1 <= 1e-9 or m2 <= 1e-9:
        return True  # 영길이 구간 — 병합해도 형상·길이 변화 없음
    cos = (v1[0] * v2[0] + v1[1] * v2[1] + v1[2] * v2[2]) / (m1 * m2)
    cos = max(-1.0, min(1.0, cos))
    return math.degrees(math.acos(cos)) >= 180.0 - tol_deg


def simplify_passthrough_nodes(net: CommonNetwork, *,
                               angle_tol_deg: float = 2.0) -> int:
    """일직선 통과하는 degree-2 base 노드를 병합. 제거된 노드 수 반환 (in-place)."""
    # 인접: nid -> 접속 배관 id 집합
    adj: dict[str, set[str]] = {}
    for pid, cp in net.pipes.items():
        adj.setdefault(cp.start, set()).add(pid)
        adj.setdefault(cp.end, set()).add(pid)

    work: list[str] = [nid for nid, cn in net.nodes.items() if cn.kind == "base"]
    removed = 0
    while work:
        nid = work.pop()
        cn = net.nodes.get(nid)
        if cn is None or cn.kind != "base":
            continue
        pids = adj.get(nid)
        if not pids or len(pids) != 2:
            continue
        pa_id, pb_id = tuple(pids)
        if pa_id == pb_id:
            continue  # 자기루프
        pa, pb = net.pipes.get(pa_id), net.pipes.get(pb_id)
        if pa is None or pb is None:
            continue
        a = pa.start if pa.end == nid else pa.end
        b = pb.start if pb.end == nid else pb.end
        if a == nid or b == nid or a == b:
            continue  # 자기루프 / 평행 생성 방지
        # ★ 헤드/노즐 직결 통과노드(드롭 엘보)는 병합 금지 — 수직 드롭 보존.
        #   하향식 헤드의 수직 드롭은 **표시 z(display_z)** 로만 존재하고 elevation_m
        #   (수리 표고)에선 헤드가 평면에 얹혀 영길이로 보인다. 헤드 x,y 를 엘보에
        #   스냅(대조 서버 _collapse_riser_to_column)하면 엘보↔헤드 구간이 영길이가 돼
        #   _collinear_3d 가 collapse → 헤드가 가지 본관 노드로 재연결되며 대각선이
        #   부활한다. 드롭 엘보를 남겨야 헤드가 엘보 바로 밑(완전 수직)에 매달린다.
        _ka = net.nodes[a].kind if a in net.nodes else ""
        _kb = net.nodes[b].kind if b in net.nodes else ""
        if _ka in ("head", "nozzle") or _kb in ("head", "nozzle"):
            continue
        # 수리 호환 — 호칭경·재질·C 동일해야 무손실
        if pa.nominal_mm != pb.nominal_mm:
            continue
        if (pa.pipe_type_label or "") != (pb.pipe_type_label or ""):
            continue
        if abs(pa.c_factor - pb.c_factor) > 1e-6:
            continue
        na, nb = net.nodes.get(a), net.nodes.get(b)
        if na is None or nb is None:
            continue
        if not _collinear_3d(na, cn, nb, angle_tol_deg):
            continue
        # 병합: pa 를 a→b 로 확장, pb·n 제거
        pa.start, pa.end = a, b
        pa.length_m += pb.length_m
        pa.equivalent_length_m += pb.equivalent_length_m
        if pb.fittings:
            pa.fittings = list(pa.fittings) + list(pb.fittings)
        del net.pipes[pb_id]
        del net.nodes[nid]
        adj.pop(nid, None)
        adj[b].discard(pb_id)
        adj[b].add(pa_id)
        removed += 1
        # 이웃이 새로 일직선 통과점이 됐을 수 있음 → 재검사
        work.append(a)
        work.append(b)
    return removed


# ────────────────────────────────────────────────────────────────────────────
# KFP 파서 — JSON 로드 → CommonNetwork
# ────────────────────────────────────────────────────────────────────────────


def parse_kfp(path: str | Path) -> CommonNetwork:
    """KFP 파일 (JSON) → CommonNetwork."""
    p = Path(path)
    # UTF-8 BOM 또는 일반 UTF-8 모두 처리
    text = p.read_text(encoding="utf-8-sig")
    data = json.loads(text)
    return kfp_dict_to_network(data)


def kfp_dict_to_network(data: dict) -> CommonNetwork:
    """KFP dict → CommonNetwork."""
    net = CommonNetwork(source_format="kfp")

    # ── 노드 (nodes + nodes_meta 통합)
    # 2.2-NODE-SCHEMA: nodes(좌표) + nodes_meta(속성) 분리.
    # 4.0-NFPA13-EQL: nodes/nodes_meta 없이 nodes_meta_runtime 단일(좌표는 meta["coords"]).
    # 두 포맷 모두 지원.
    metas = data.get("nodes_meta") or data.get("nodes_meta_runtime", {})
    coords = data.get("nodes")
    if not coords:
        coords = {nid: m.get("coords", [0.0, 0.0, 0.0]) for nid, m in metas.items()}
    for nid, xyz in coords.items():
        meta = metas.get(nid, {})
        # KFP type_id 종류 (17 파일 분석):
        #   base   — 분기/끝점
        #   nozzle — 스프링클러/분무 노즐 (k_factor_si 보유)
        #   head   — 스프링클러 헤드 (k_factor_si + head_spec_name 보유) ★ nozzle 와 같이 outlet
        #   wt     — 수원 (Water Tank, water_level 또는 required_pressure_bar)
        #   pump   — 펌프 (rated_q/p, shutoff_p, peak_q/p)
        #   valve  — 일반 밸브 (fitting_id)
        #   prv    — PRV 감압 밸브 (pressure_setting_bar + loss_coefficient)
        #   orifice — 오리피스 (hole_diameter + orifice_discharge_coeff)
        cn = CommonNode(
            id=nid,
            x=float(xyz[0]) if len(xyz) > 0 else 0.0,
            y=float(xyz[1]) if len(xyz) > 1 else 0.0,
            elevation_m=float(meta.get("elevation_m", xyz[2] if len(xyz) > 2 else 0.0)),
            kind=str(meta.get("type_id", "base")).lower() or "base",
            k_factor_si=meta.get("k_factor_si"),
            pressure_bar=None,
            valve_type=None,
            is_check_valve=bool(meta.get("is_check_valve", False)),
            raw={k: v for k, v in meta.items() if k not in (
                "id", "coords", "type", "type_id", "elevation_m",
                "is_active", "k_factor_si", "is_check_valve",
            )},
        )
        # 표시 전용 z (coords[2]) 보존 — 라이저 기둥·헤드 상/하향 돌출이 여기 담긴다.
        # 수리 elevation_m 과 분리. geometry 변환이 display 우선으로 z 를 그린다.
        if len(xyz) > 2:
            cn.raw["display_z_m"] = float(xyz[2])
        # 수원 (wt) — boundary pressure 결정 우선순위:
        # 1) required_pressure_bar 가 0 초과 → 그대로
        # 2) water_level (m) 이 0 초과 → 정수두 ρgh = m × 9.81/100 bar
        # 3) fallback → 1.013 bar (대기압)
        if cn.kind == "wt":
            req_p = meta.get("required_pressure_bar") or 0
            wl = meta.get("water_level") or 0
            if req_p > 0:
                cn.pressure_bar = float(req_p)
            elif wl > 0:
                cn.pressure_bar = float(wl) * 0.0981  # m H₂O → bar
            else:
                cn.pressure_bar = 1.013  # 대기압
        # pump — boundary pressure = rated_p (펌프 outlet 압력)
        if cn.kind == "pump":
            cn.pump_curve = {
                "rated_q": meta.get("rated_q", 0.0),
                "rated_p": meta.get("rated_p", 0.0),
                "shutoff_p": meta.get("shutoff_p", 0.0),
                "peak_q": meta.get("peak_q", 0.0),
                "peak_p": meta.get("peak_p", 0.0),
                "curve_data": meta.get("pump_curve_data", []),
            }
            # pump 노드도 시스템 source — boundary pressure 필요
            rp = meta.get("rated_p") or 0
            if rp > 0:
                cn.pressure_bar = float(rp)
            else:
                cn.pressure_bar = 1.013
        # valve — type_id "valve" + fitting_id 보조
        if cn.kind == "valve":
            cn.valve_type = meta.get("fitting_id") or "gate_valve"
        # PRV — pressure_setting_bar
        if cn.kind == "prv":
            cn.valve_type = "prv"
            ps = meta.get("pressure_setting_bar")
            if ps:
                cn.pressure_bar = float(ps)
        # orifice — hole_diameter + discharge_coeff (별도 처리)
        if cn.kind == "orifice":
            cn.valve_type = "orifice"
        net.nodes[nid] = cn

    # ── 파이프
    pipe_data = data.get("pipe_data", {})
    for pid, pdat in pipe_data.items():
        fittings = []
        for f in pdat.get("fittings", []) or []:
            # KFP fitting 표현은 두 가지: (a) string ("Elbow", "Tee") — 사용자 도면에서
            # 관찰된 형태, (b) dict {"id":..., "count":...} — 확장 schema. 둘 다 처리.
            if isinstance(f, str):
                ftype = f.lower()
                fcount = 1
                lod = None
            elif isinstance(f, dict):
                ftype = str(f.get("type_id") or f.get("id") or "").lower()
                fcount = int(f.get("count", 1))
                lod = f.get("L_over_D")
            else:
                continue
            # 표준화 — KFP 의 elbow_45 / "Elbow 45" → SDF elbow-45 형태
            ftype_std = ftype.replace("_", "-").replace(" ", "-")
            fittings.append(CommonFitting(
                type_id=ftype_std, count=fcount, l_over_d=lod,
            ))
        cp = CommonPipe(
            id=pid,
            start=pdat.get("start", ""),
            end=pdat.get("end", ""),
            length_m=float(pdat.get("length_m", 0.0)),
            diameter_inner_mm=float(pdat.get("diameter", 0.0)),
            nominal_mm=int(pdat.get("nominal_mm", 0)),
            c_factor=float(pdat.get("C", 120.0)),
            roughness_mm=float(pdat.get("roughness_mm", 0.15)),
            pipe_type_label=str(pdat.get("type", "KSD 3507")),
            fittings=fittings,
            equivalent_length_m=float(pdat.get("equivalent_length", 0.0)),
            raw={k: v for k, v in pdat.items() if k not in (
                "start", "end", "length_m", "diameter", "nominal_mm",
                "C", "roughness_mm", "type", "fittings", "equivalent_length",
            )},
        )
        net.pipes[pid] = cp

    # ── 라이브러리 + 설정
    nlib = data.get("nozzle_library", {}) or {}
    for cat in nlib.get("categories", []):
        for item in cat.get("items", []):
            net.nozzle_library.append({
                "id": item.get("id"),
                "display_name": item.get("display_name"),
                "K_SI": item.get("K_SI") or item.get("K_val"),
                "K_US": item.get("K_US"),
                "category": cat.get("category_id"),
            })
    flib = data.get("fittings_library", {}) or {}
    for cat in flib.get("categories", []):
        for item in cat.get("items", []):
            net.fitting_library.append({
                "id": item.get("id"),
                "display_name": item.get("display_name"),
                "type_id": item.get("type_id"),
                "L_over_D": item.get("L_over_D"),
                "category": cat.get("category_id"),
                "is_check_valve": item.get("is_check_valve", False),
            })
    # 4.0-NFPA13-EQL: fittings_library_v3 (DN별 등가길이 표). round-trip 은 배관별
    # equivalent_length 를 직접 읽으므로 카탈로그는 참조용으로만 보존.
    for item in data.get("fittings_library_v3", []) or []:
        net.fitting_library.append({
            "id": item.get("id"),
            "display_name": item.get("name_en") or item.get("name_ko"),
            "type_id": item.get("category"),
            "L_over_D": None,
            "equivalent_lengths": item.get("equivalent_lengths"),
            "category": item.get("category"),
            "is_check_valve": "check" in (item.get("category") or "").lower(),
        })
    plib = data.get("pump_library", {}) or {}
    for pump in plib.get("pumps", []) or []:
        net.pump_library.append(dict(pump))
    net.design_settings = dict(data.get("design_settings", {}))
    net.project_meta = dict(data.get("report_common", {}))
    net.project_meta["_source_version"] = data.get("version", "")
    return net


# ────────────────────────────────────────────────────────────────────────────
# 표시좌표 스케일 — 등각 그리드 포맷(KFP/HAS) 공용
#   통합망 schematic 좌표는 라이저(2m)+헤드망이 3×3m 안에 압축돼, 등각 그리드 위에
#   배관망이 점처럼 작게 보인다. 노드 간 중앙값 최근접거리를 target_spacing 에 맞춰
#   균일 확대해 레퍼런스(.kfp/.has)급 크기로 펼친다. 위상·형상·수리계산 불변
#   (KFP 길이=pipe length_m, HAS 길이=PipeLength 별도 저장 → 좌표는 표시 전용).
# ────────────────────────────────────────────────────────────────────────────


def _coord_scale(nodes: list["CommonNode"]) -> float:
    """좌표를 미터로 정규화하는 배율. bbox span 이 수 km 면 mm 로 보고 0.001."""
    xs = [n.x for n in nodes]
    ys = [n.y for n in nodes]
    if not xs:
        return 1.0
    span = max(max(xs) - min(xs), max(ys) - min(ys))
    return 0.001 if span > 5000.0 else 1.0


def _spread_factor(
    nodes: list["CommonNode"],
    base_scale: float,
    *,
    target_spacing_m: float = 0.7,
    cap: float = 80.0,
) -> float:
    """뭉친 schematic 을 펼치는 추가 배율 — 노드 간 중앙값 최근접거리를 target 에 맞춤.

    이미 충분히 펼쳐진 망(중앙값 ≥ target)은 1.0. 위상·형상 보존, 수리계산 무영향.
    """
    pts = [
        (n.x * base_scale, n.y * base_scale)
        for n in nodes
        if math.isfinite(n.x) and math.isfinite(n.y)
    ]
    if len(pts) < 4:
        return 1.0
    # 대형 망은 표본으로 중앙값 추정(O(sample×N)). 145급은 전수.
    if len(pts) > 600:
        import random

        sample = random.Random(0).sample(pts, 600)
    else:
        sample = pts
    nn: list[float] = []
    for xi, yi in sample:
        best = math.inf
        for xj, yj in pts:
            if xj == xi and yj == yi:
                continue
            d = math.hypot(xi - xj, yi - yj)
            if d < best:
                best = d
        if math.isfinite(best) and best > 0:
            nn.append(best)
    if not nn:
        return 1.0
    nn.sort()
    median_nn = nn[len(nn) // 2]
    if median_nn <= 0 or median_nn >= target_spacing_m:
        return 1.0
    return min(cap, target_spacing_m / median_nn)


def _display_xy_scale(nodes: list["CommonNode"]) -> float:
    """등각 그리드 포맷용 표시좌표 배율(= 미터정규화 × 펼침). x,y 에만 적용."""
    bs = _coord_scale(nodes)
    return bs * _spread_factor(nodes, bs)


# ────────────────────────────────────────────────────────────────────────────
# KFP 직렬화 — CommonNetwork → JSON
# ────────────────────────────────────────────────────────────────────────────


def emit_kfp(net: CommonNetwork, path: str | Path | None = None,
             *, coord_scale: float = 1.0, display_geometry: bool = False) -> dict:
    """CommonNetwork → KFP dict (JSON 직렬화 가능). path 주어지면 파일에도 쓰기.

    ``coord_scale`` — (API 호환용, 무시됨) 과거 표시좌표 확대 배율. K-Fire_Solver 가
    노드 3D 좌표거리에서 배관장을 역산하므로(좌표거리 == length_m), 좌표에 임의
    배율을 곱하면 length_m 과 어긋나 배관장이 왜곡된다. 좌표는 length_m 에 맞춰
    자가보정(실축척 미터)되며, 사용자 배율은 적용하지 않는다.

    ``display_geometry`` — 통합망(combined) 전용. False(단독 도면)면 좌표 z=실표고
    + length_m 자가보정으로 좌표거리==length_m(수리 정확). True 면 미리보기와 동일한
    스키매틱 표시좌표(라이저=수직 기둥, display_z_m)를 쓰고 균일 확대로 노드를
    레퍼런스급 크기로 펼친다(노드 상대 축소, 평면/계통도 비율은 미리보기와 일치).
    이때 좌표는 표시 전용이고 length_m·elevation_m 는 실값(수리 권위값)으로 보존되나,
    K-Fire_Solver 가 좌표거리에서 배관장을 역산하면 통합망 표/자동수리는 스키매틱이
    된다 → 정밀 수리는 단독 도면 KFP 를 쓸 것.

    K-Fire_Solver 호환성:
    - type 필드는 라이브러리 표준 영문명 사용 ("Water Spray", "WT", "Pump" 등).
      한글 generic 라벨 ("노즐", "수원") 은 K-Fire_Solver 가 못 알아봄.
    - has_check_valve / is_check_valve 기본 True (전체 샘플 KFP 일관).
    - 내부 추적용 key (sdf_label, io_node, sdf_lib_item) 는 출력에서 제거 —
      K-Fire_Solver 의 strict validator 가 unknown key 만나면 node drop →
      Error 223 "not enough nodes" 발생 가능성 차단.
    - license_tag "TRIAL" (전체 샘플 일관, "CONVERTED" 는 비표준).
    """
    # type 라벨 매핑 — K-solver 가 type 필드로 노드 종류 식별. 17 샘플 분석:
    #   base   → "기본"            (Korean OK, 모든 샘플)
    #   nozzle → "Water Spray"     (분무 노즐 표준 라벨)
    #   head   → "Head"            (스프링클러 헤드)
    #   wt     → "WT"              (Water Tank)
    #   valve  → valve_type 별 ("Alarm Valve", "Preaction Valve", "Gate Valve" 등)
    #   pump   → "Pump"
    #   prv    → "Prv"
    #   orifice→ "Orifice"
    TYPE_LABEL_BY_KIND = {
        "base": "기본",
        "nozzle": "Water Spray",
        "head": "Head",
        "wt": "WT",
        "pump": "Pump",
        "prv": "Prv",
        "orifice": "Orifice",
    }
    VALVE_TYPE_LABELS = {
        "alarm_valve": "Alarm Valve",
        "preaction_valve": "Preaction Valve",
        "dry_pipe_valve": "Dry Pipe Valve",
        "gate_valve": "Gate Valve",
        "globe_valve": "Globe Valve",
        "ball_valve": "Ball Valve",
        "butterfly_valve": "Butterfly Valve",
        "angle_valve": "Angle Valve",
        "swing_check": "Swing Check",
        "smolensky_check": "Smolensky Check",
    }

    def _type_label_for(cn) -> str:
        if cn.kind == "valve":
            vt = (cn.valve_type or "gate_valve").lower().replace("-", "_")
            return VALVE_TYPE_LABELS.get(vt, "Gate Valve")
        return TYPE_LABEL_BY_KIND.get(cn.kind, "기본")

    # K-solver 가 unknown key 만나면 node drop 가능 → 우리 내부 추적용 키는 제외.
    # pump_output 은 우리 내부 추적용(HAS pumpBlk ENodeId / KFP 토출 파이프 합성).
    # 참조 KFP 의 펌프 meta 엔 이 키가 없으므로 K-solver 노출 meta 에서 제외.
    # display_z_m 은 coords[2] 와 동치(파싱 시 line 289 에서 재구성) — 4.0 템플릿
    # nodes_meta_runtime 에는 없으므로 출력 meta 에서 제외해 strict validator 와 정합.
    INTERNAL_RAW_KEYS = {"sdf_label", "io_node", "sdf_lib_item", "pump_name", "pump_output", "display_z_m"}

    nodes = {}
    nodes_meta = {}
    node_types = {}
    for nid, cn in net.nodes.items():
        # ★ 단독 도면: KFP 좌표 z = 실표고(elevation_m). K-Fire_Solver 는 노드 3D
        # 좌표거리에서 배관장을 역산하므로(레퍼런스 .kfp 검증: 좌표거리 == length_m),
        # 표시 전용 display_z 를 쓰면 length_m 과 어긋나 70m·350m 같은 비현실 배관장이
        # 나온다. ★ 통합망(display_geometry): 미리보기와 동일한 스키매틱 표시 z
        # (라이저=기둥, display_z_m)를 써 비율을 미리보기와 일치시킨다(아래 자가보정 분기
        # 에서 length 보정 대신 균일 확대). elevation_m 는 양쪽 모두 실표고 보존.
        disp_z = (cn.raw.get("display_z_m", cn.elevation_m)
                  if display_geometry else cn.elevation_m)
        nodes[nid] = [cn.x, cn.y, disp_z]
        type_label = _type_label_for(cn)
        node_types[nid] = type_label
        # required_pressure_bar 결정:
        # - wt: 보통 0.0 (water_level + elevation 으로 정수두 직접 계산). pressure_bar
        #   에 fallback 값 (water_level × 0.0981) 이 들어 있으면 무시.
        # - nozzle/head: 최소 1 bar 보장 (NFTC, K-solver 의 노즐 비활성화 회피).
        # - 그 외: cn.pressure_bar 또는 0.
        if cn.kind == "wt":
            # water_level 이 있으면 required_pressure_bar 는 0 (이중 계산 회피).
            wl_raw = cn.raw.get("water_level")
            req_p = 0.0 if (wl_raw and wl_raw > 0) else (cn.pressure_bar or 0.0)
        elif cn.kind in ("nozzle", "head"):
            req_p = cn.pressure_bar or 1.0
            if req_p == 0.0:
                req_p = 1.0
        else:
            req_p = cn.pressure_bar or 0.0
        meta = {
            "id": nid,
            "coords": [cn.x, cn.y, disp_z],  # 표시 전용 (elevation_m 과 분리)
            "type": type_label,
            "type_id": cn.kind,
            "elevation_m": cn.elevation_m,  # 수리계산용 실표고 (표시 z 와 독립)
            "is_active": True,
            "is_closed": False,
            # ★ is_check_valve = "이 노드 자체가 체크밸브인가" — valve 중 일부 (9/38)
            # 만 True. base/nozzle/wt 는 전부 False (전체 샘플 확인).
            "is_check_valve": bool(cn.is_check_valve) if cn.kind == "valve" else False,
            # k_factor_si float 정밀도 drift 방지 — K [L/min·bar^-0.5] 는 보통
            # 3 자릿수 ("24.1", "80.0"). SDF round-trip 시 24.1 → 0.0004017 m³/s
            # → 24.100019999 처럼 부동소수 노이즈. K-solver 가 round 비교하면 다른
            # 노즐로 인식할 수도 있어 소수 4자리 round.
            "k_factor_si": round(cn.k_factor_si, 4) if cn.k_factor_si else cn.k_factor_si,
            "required_pressure_bar": req_p,
            "base_demand_lps": 0.0,
            "fitting_id": (VALVE_TYPE_LABELS.get((cn.valve_type or "").lower().replace("-", "_"), "VALVE_GATE").upper().replace(" ", "_")
                            if cn.kind == "valve" else None),
            "head_spec_name": None,
            "nozzle_name": cn.raw.get("nozzle_name") or (
                "45L/min" if cn.kind == "nozzle" else None),
            # ★ wt(고가수조)는 펌프가 아님 — PUMP_001(전제로 rated=0 인 Default Pump)을
            # 달면 K-solver 가 shutoff=0 인 닫힌 펌프로 보고 유량을 0 으로 막음.
            # 참조 .kfp(시상수도연결)의 WT 도 pump_library_id=null. wt 정수두는
            # elevation(+water_level)으로 공급되므로 라이브러리 펌프 불필요.
            "pump_library_id": "PUMP_001" if cn.kind == "pump" else None,
            "rated_q": (cn.pump_curve or {}).get("rated_q", 0.0),
            "rated_p": (cn.pump_curve or {}).get("rated_p", 0.0),
            "shutoff_p": (cn.pump_curve or {}).get("shutoff_p", 0.0),
            "peak_q": (cn.pump_curve or {}).get("peak_q", 0.0),
            "peak_p": (cn.pump_curve or {}).get("peak_p", 0.0),
            "water_level": 0.0,
            "pump_curve_data": (cn.pump_curve or {}).get("curve_data", []),
            "has_check_valve": True,   # ★ 전체 샘플 True
            "pressure_setting_bar": None,
            "loss_coefficient": 0.0,
            "hole_diameter": 0.0,
            "orifice_discharge_coeff": 0.65,  # ★ 4.0 기본값 (orifice 노드만 의미 — 물리 표준 0.65)
            "pipe_dn": 0.0,
            "check_valve_direction": "",
            # ── 4.0-NFPA13-EQL 추가 필드 (전체 노드 공통, 4.0 템플릿 기본값) ──
            "category_id": "",
            "hose_stream_flow_lps": 6.666666666666667,
            "fdt_differential_ratio": 5.5,
            "fdt_primary_water_gauge_bar": 5.0,
            "fdt_p_initial_gauge_bar": 2.4,
            "fdt_gas": "air",
            "fdt_temp_c": 20.0,
            "fdt_trip_pressure_gauge_bar": 0.0,
            "fdt_interlock_mode": "single",
            "fdt_accel_installed": True,
            "fdt_accel_kind": "mechanical",
            "fdt_accel_drop_bar": 0.3,
            "fdt_accel_trip_time_s": 3.0,
        }
        # 원본 raw 메타로 덮어쓰기 (round-trip 보존). 단 내부 추적용 키 제외.
        meta.update({k: v for k, v in cn.raw.items() if k not in INTERNAL_RAW_KEYS})
        # type 필드 - meta.update 후 raw 의 한글 type 이 덮어쓸 수 있으니 마지막 강제.
        meta["type"] = type_label
        node_types[nid] = type_label
        # ★ 고가수조 (wt, 펌프 없음) — K-solver 가 driving head 0 으로 보고 계산
        # 거부. water_level 이 raw 에 없으면 boundary pressure 또는 elevation 으로
        # 정수두 환산. K-solver wt 정수두 = elevation + water_level.
        if cn.kind == "wt" and not meta.get("water_level"):
            if cn.pressure_bar and cn.pressure_bar > 0:
                meta["water_level"] = round(cn.pressure_bar * 10.197, 3)
            elif cn.elevation_m and cn.elevation_m > 0:
                meta["water_level"] = 2.0
        nodes_meta[nid] = meta

    # 공통 standard 추정 — 첫 pipe 의 type 로 (각 pipe 마다 다를 수도 있지만 보통 통일)
    pipe_standard_default = "KSD3507 (SPP)"
    if net.pipes:
        first_pipe = next(iter(net.pipes.values()))
        pipe_standard_default = first_pipe.pipe_type_label or pipe_standard_default

    pipe_data = {}
    for pid, cp in net.pipes.items():
        fittings_out = []
        eq_length_total = 0.0
        # inner_d_mm lookup — nominal 만 있으면 inner 채우기. 권위 SLF 우선:
        # FX/DP 처럼 inner==nominal(25.0) 인 스케줄도 .slf 권위값 25.0 을 보존하고
        # K-solver 강관 표준(27.5)으로 뭉개지 않는다. .slf 미매칭 시 K-solver fallback.
        raw_type = cp.raw.get("sdf_pipe_type") or cp.pipe_type_label
        inner_mm = cp.diameter_inner_mm
        if inner_mm <= 0 or abs(inner_mm - cp.nominal_mm) < 0.5:
            slf_inner = _slf_inner_d_mm(cp.nominal_mm, raw_type)
            if slf_inner is None:
                slf_inner = _slf_inner_d_mm(cp.nominal_mm, cp.pipe_type_label or pipe_standard_default)
            inner_mm = (slf_inner if slf_inner is not None
                        else _lookup_inner_d_mm(cp.nominal_mm, cp.pipe_type_label or pipe_standard_default))
        # 조도도 .slf 권위값으로 정합 (매칭 시). 미매칭이면 파싱값 유지.
        slf_rough = _slf_roughness_mm(raw_type)
        if slf_rough is None:
            slf_rough = _slf_roughness_mm(cp.pipe_type_label or pipe_standard_default)
        roughness_out = slf_rough if slf_rough is not None else cp.roughness_mm
        # fitting display 명 + L_over_D 계산 → equivalent_length
        # K-solver reference KFP 의 표기 규약 (1-1.업무시설_201동_28F 분석):
        #   tee-branch / tee_branch → "Tee(Branch)"  ← 분기 tee
        #   elbow → "Elbow"
        #   elbow-45 → "Elbow 45"
        #   외 종류 → snake_case → Title Case
        FITTING_DISPLAY_MAP = {
            "tee-branch": "Tee(Branch)",
            "tee_branch": "Tee(Branch)",
            "tee(branch)": "Tee(Branch)",
            "elbow": "Elbow",
            "tee": "Tee(Run)",   # Tee through-flow (자주 안 나옴)
            "butterfly": "Butterfly Valve",
            "check": "Swing Check Valve",
        }
        for f in cp.fittings:
            display = FITTING_DISPLAY_MAP.get(
                f.type_id.lower(),
                f.type_id.replace("-", " ").replace("_", " ").title(),
            )
            lod = f.l_over_d or _lookup_fitting_lod(display)
            if lod is not None and inner_mm > 0:
                # 등가길이 = (L/D) × inner_d (m 단위)
                eq_length_total += (lod * inner_mm / 1000.0) * max(1, f.count)
            for _ in range(max(1, f.count)):
                fittings_out.append(display)
        # 우선순위: 명시된 equivalent_length_m > 계산된 eq_length_total
        eq_len = cp.equivalent_length_m if cp.equivalent_length_m > 0 else eq_length_total
        pdat = {
            "start": cp.start,
            "end": cp.end,
            "type": cp.pipe_type_label or pipe_standard_default,
            "diameter": inner_mm,             # ★ inner (K-solver hydraulic 계산용)
            "nominal_mm": cp.nominal_mm,
            "length_m": cp.length_m,
            "equivalent_length": round(eq_len, 5),  # ★ 자동 계산
            "C": cp.c_factor,
            "roughness_mm": roughness_out,
            "fittings": fittings_out,
            "flow_lpm": 0.0,
            "velocity_mps": 0.0,
            "headloss_m": 0.0,
        }
        # K-solver 가 unknown key 에 민감 — 내부 추적용 sdf_label / sdf_fittings 제외
        pdat.update({k: v for k, v in cp.raw.items() if k not in ("sdf_label", "sdf_fittings", "sdf_pipe_type")})
        pipe_data[pid] = pdat

    if display_geometry:
        # ★ 통합망 — 좌표는 미리보기와 동일한 스키매틱 표시(라이저=기둥, display_z).
        # length_m·elevation_m 는 실값(수리 권위값)으로 보존. 균일 확대(x,y,z 동일 배율)
        # 로 노드를 레퍼런스급 크기로 펼쳐(노드 상대 축소) 평면/계통도 비율은 미리보기와
        # 일치시킨다. 균일이라 비율 불변, length 자가보정은 하지 않는다(표시 전용).
        s_disp = _display_xy_scale(list(net.nodes.values()))
        if s_disp > 0 and abs(s_disp - 1.0) > 1e-9:
            for _xyz in nodes.values():
                _xyz[0] *= s_disp
                _xyz[1] *= s_disp
                _xyz[2] *= s_disp
            for _m in nodes_meta.values():
                _c = _m.get("coords")
                if _c:
                    _c[0] *= s_disp
                    _c[1] *= s_disp
                    if len(_c) > 2:
                        _c[2] *= s_disp
    else:
        # ★ 단독 도면 좌표 자가보정 — K-Fire_Solver 는 노드 3D 좌표거리에서 배관장을
        # 역산한다(레퍼런스 .kfp: 좌표거리 == length_m). schematic x,y 는 임의 단위
        # (emit_sdf 가 최장변 3000 으로 정규화)라 그대로 두면 length_m 과 어긋나 비현실
        # 배관장이 나온다. 수평에 가까운 배관에서 length_m/2D거리 의 중앙값을 배율로 잡아
        # x,y 를 미터로 환산하면, z=실표고(미터)와 합쳐 좌표거리 == length_m 이 성립한다
        # (펌프 토출관 합성 전에 적용해 합성길이도 보정좌표로 계산되게 함).
        _ratios: list[float] = []
        for _p in pipe_data.values():
            _s, _e = nodes.get(_p["start"]), nodes.get(_p["end"])
            if not _s or not _e:
                continue
            _d2 = math.hypot(_s[0] - _e[0], _s[1] - _e[1])
            _len_m = _p.get("length_m") or 0.0
            if _d2 <= 1e-9 or _len_m <= 0:
                continue
            _dz = abs((_s[2] or 0.0) - (_e[2] or 0.0))
            if _dz <= 0.25 * _len_m:  # 수평에 가까운 배관만 (z 성분이 작아야 2D≈length)
                _ratios.append(_len_m / _d2)
        if _ratios:
            _ratios.sort()
            s_xy = _ratios[len(_ratios) // 2]
        else:
            s_xy = _coord_scale(list(net.nodes.values()))
        if s_xy > 0 and abs(s_xy - 1.0) > 1e-9:
            for _xyz in nodes.values():
                _xyz[0] *= s_xy
                _xyz[1] *= s_xy
            for _m in nodes_meta.values():
                _c = _m.get("coords")
                if _c:
                    _c[0] *= s_xy
                    _c[1] *= s_xy

    # 노드/파이프 counter 추정 (다음 ID 할당용)
    max_node_n = 0
    for nid in nodes:
        if nid.startswith("N"):
            try:
                max_node_n = max(max_node_n, int(nid[1:]))
            except ValueError:
                pass
    max_pipe_n = 0
    for pid in pipe_data:
        if pid.startswith("P"):
            try:
                max_pipe_n = max(max_pipe_n, int(pid[1:]))
            except ValueError:
                pass

    # ★ 펌프 토출 파이프 합성 — PIPENET 의 Pump-fan 은 펌프 자체가 input→output 링크라
    # net.pipes 에 안 들어온다(Links>Pipe-set>Pipe 만 파이프로 파싱). K-Fire_Solver
    # KFP 는 HAS 와 달리 펌프-엣지 개념이 없어 펌프를 pipe_data 로 연결된 인라인 노드로
    # 요구한다 — 참조 KFP 의 모든 펌프(2zone booster N1/N113, 시상수도연결 N9,
    # 성능시험 N1)는 토출 파이프로 연결돼 있고 isolated 노드가 0 이다. 토출 파이프가
    # 없으면 펌프 노드가 고립돼 "특정 노드 누락" 으로 계산이 막힌다. parse_sdf 가
    # 보존한 pump_output 으로 토출 파이프를 합성해 pipe_data 에만 추가(net.pipes 는
    # 불변 → HAS pumpBlk 의 별도 펌프-엣지와 중복 pline 안 생김).
    _next_pnum = max_pipe_n + 1
    for nid, cn in net.nodes.items():
        if cn.kind != "pump":
            continue
        out_ref = cn.raw.get("pump_output")
        if not out_ref or str(out_ref) not in nodes or nid not in nodes:
            continue
        out_ref = str(out_ref)
        if any((p["start"] == nid and p["end"] == out_ref) or
               (p["start"] == out_ref and p["end"] == nid)
               for p in pipe_data.values()):
            continue
        # 토출 노드에 닿는 대표 파이프에서 구경/타입/C 복사 (없으면 100A 기본)
        ref_pipe = next((p for p in pipe_data.values()
                          if p["start"] == out_ref or p["end"] == out_ref),
                         next(iter(pipe_data.values()), None)) or {}
        sx, sy, sz = nodes[nid]
        ex, ey, ez = nodes[out_ref]
        _len = math.sqrt((sx - ex) ** 2 + (sy - ey) ** 2
                          + ((sz or 0.0) - (ez or 0.0)) ** 2)
        new_pid = f"P{_next_pnum}"
        _next_pnum += 1
        pipe_data[new_pid] = {
            "start": nid,
            "end": out_ref,
            "type": ref_pipe.get("type", pipe_standard_default),
            "diameter": ref_pipe.get("diameter", 105.3),
            "nominal_mm": ref_pipe.get("nominal_mm", 100),
            "length_m": round(_len, 5) if _len > 0 else 1.0,
            "equivalent_length": 0.0,
            "C": ref_pipe.get("C", 120.0),
            "roughness_mm": ref_pipe.get("roughness_mm", 0.15),
            "fittings": [],
            "flow_lpm": 0.0,
            "velocity_mps": 0.0,
            "headloss_m": 0.0,
        }
    max_pipe_n = _next_pnum - 1

    # ★ 라이브러리 카탈로그 — 권위 SLF(2. Pipenet_hand.slf) 우선, 없으면 K-solver fallback.
    _slf_pipe_lib = _slf_pipe_library_for_kfp()
    _slf_noz_lib = _slf_nozzle_library_for_kfp()
    _slf_pump_list = _slf_pump_library_for_kfp()
    if _slf_noz_lib is not None:
        # .slf 는 head 만 정의(SP-HEAD/INDOOR HYDRANT). spray nozzle 정의는 .slf 에
        # 없으므로 K-solver spray 카테고리를 그대로 덧붙여 분사 헤드 매칭을 유지한다.
        _nozzle_library = {
            "schema_version": "2.0",
            "categories": list(_slf_noz_lib["categories"]) + [
                {
                    "category_id": "nozzle",
                    "display_name": "Nozzle",
                    "is_system_protected": True,
                    "type_id": "nozzle",
                    "items": [
                        {
                            "id": f"NOZZLE_{n['name'].replace(' ', '_').upper()}",
                            "display_name": n["name"],
                            "K_val": n["K_val"],
                            "Q_lpm": n["Q_lpm"],
                            "P_bar": n["P_bar"],
                            "min_p": n["min_p"],
                            "type": "spray",
                            "is_system_protected": True,
                        }
                        for n in KSOLVER_NOZZLE_LIBRARY["nozzles"]
                    ],
                },
            ],
        }
    else:
        _nozzle_library = {
            "schema_version": "2.0",
            "categories": [
                {
                    "category_id": "head",
                    "display_name": "Head",
                    "is_system_protected": True,
                    "type_id": "head",
                    "items": [
                        {
                            "id": f"HEAD_{int(n['K_val'])}",
                            "display_name": n["name"],
                            "K_SI": n["K_val"],
                            "K_US": n["K_val"] / 1.4159,
                            "is_system_protected": True,
                        }
                        for n in KSOLVER_NOZZLE_LIBRARY["nozzles"]
                    ],
                },
                {
                    "category_id": "nozzle",
                    "display_name": "Nozzle",
                    "is_system_protected": True,
                    "type_id": "nozzle",
                    "items": [
                        {
                            "id": f"NOZZLE_{n['name'].replace(' ', '_').upper()}",
                            "display_name": n["name"],
                            "K_val": n["K_val"],
                            "Q_lpm": n["Q_lpm"],
                            "P_bar": n["P_bar"],
                            "min_p": n["min_p"],
                            "type": "spray",
                            "is_system_protected": True,
                        }
                        for n in KSOLVER_NOZZLE_LIBRARY["nozzles"]
                    ],
                },
            ],
        }

    out = {
        # ★ 4.0-NFPA13-EQL: 노드는 nodes_meta_runtime 단일 저장. 2.2 의 nodes(좌표)/
        # node_types/nodes_meta 3중 미러는 4.0 이 읽지 않고 strict validator 가 거부할
        # 수 있어 4.0 템플릿(ewqwqe.kfp)과 동일하게 제거. 좌표는 각 노드 meta["coords"].
        "pipe_data": pipe_data,
        "pipe_id_counter": max_pipe_n + 1,
        "node_counter": {"N": max_node_n + 1},
        "nodes_meta_runtime": nodes_meta,
        "design_settings": net.design_settings or {
            "min_required_pressure_bar": 1.0,
            "calculation_method": "H-W",
            "standard": pipe_standard_default,
            # 첫 pipe 의 C 사용 — 전체 동일 가정 (PIPENET SDF 패턴)
            "C": float(next(iter(net.pipes.values())).c_factor) if net.pipes else 120.0,
            "roughness_mm": 0.15,
            "head": "80(5.6)",
            "notes": "",
        },
        # ★ K-solver 의 PDF/리포트 생성기는 report_common 의 design_area / company /
        # date 등을 필수로 참조. 비어 있으면 KFP 로드 후 즉시 KeyError 가능성.
        "report_common": (
            {k: v for k, v in net.project_meta.items() if not k.startswith("_")}
            or {
                "design_area": "Converted Network",
                "designer": "",
                "company": "",
                "project": "Converted from SDF",
                "date": datetime.datetime.now().strftime("%Y년 %m월 %d일"),
                "notes": "",
            }
        ),
        # ★ K-Fire_Solver 표준 라이브러리 포함 — solver 가 nozzle K / pipe inner /
        # fitting L/D lookup 위해 필요. 우리 변환기는 이걸 default 로 동봉해 K-solver
        # 가 KFP 받자마자 hydraulic 계산 시작 가능하게 함.
        "pipe_library": _slf_pipe_lib or copy.deepcopy(_try_load_ksolver_libraries()["pipe"]),
        "nozzle_library": _nozzle_library,
        # ★ 4.0-NFPA13-EQL: 피팅은 NFPA13 등가길이 카탈로그(fittings_library_v3).
        # 2.2 의 L/D 기반 fittings_library 는 4.0 이 읽지 않음 → v3 로 교체.
        "fittings_library_v3": _fittings_library_v3(),
        # ★ pump_library — pump 노드의 실제 곡선 데이터로 자동 채우기
        "pump_library": {
            "pumps": (list(_slf_pump_list) + list(net.pump_library) + [
                {
                    "id": f"PUMP_{i+1:03d}",
                    "name": (cn.raw.get("pump_name") or f"Pump {i+1}"),
                    "rated_q": (cn.pump_curve or {}).get("rated_q", 0.0),
                    "rated_p": (cn.pump_curve or {}).get("rated_p", 0.0),
                    "shutoff_p": (cn.pump_curve or {}).get("shutoff_p", 0.0),
                    "peak_q": (cn.pump_curve or {}).get("peak_q", 0.0),
                    "peak_p": (cn.pump_curve or {}).get("peak_p", 0.0),
                }
                for i, (nid, cn) in enumerate(net.nodes.items())
                if cn.kind == "pump" and cn.pump_curve
            ]) or [
                {"id": "PUMP_001", "name": "Default Pump",
                 "rated_q": 0.0, "rated_p": 0.0,
                 "shutoff_p": 0.0, "peak_q": 0.0, "peak_p": 0.0},
            ],
        },
        # ★ 솔버 4.0 전용 — 구버전(2.2)은 4.0 에서 열리지 않으므로 항상 4.0 으로 출력.
        "version": "4.0-NFPA13-EQL",
        # ★ 전체 샘플 KFP 는 "TRIAL". "CONVERTED" 는 K-solver 가 미인식 → 라이센스
        # 검증 실패로 로드 거부 가능. TRIAL 통일.
        "license_tag": "TRIAL",
        "sig": "",
    }
    if path is not None:
        Path(path).write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    return out


# ────────────────────────────────────────────────────────────────────────────
# SDF 파서 — XML 로드 → CommonNetwork
# ────────────────────────────────────────────────────────────────────────────


def _strip_ns(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _find_child(elem: ET.Element | None, name: str) -> ET.Element | None:
    if elem is None:
        return None
    for c in elem:
        if _strip_ns(c.tag) == name:
            return c
    return None


def _iter_children(elem: ET.Element | None, name: str) -> list[ET.Element]:
    if elem is None:
        return []
    return [c for c in elem if _strip_ns(c.tag) == name]


def _ffloat(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


# 표준 호칭경 (mm) — KFP 와 PIPENET 공통. SDF 의 bore (inner diameter, m)
# 로부터 nominal 을 round 할 때 가장 가까운 표준 값 사용.
STANDARD_NOMINAL_MM = (15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300)


def _nearest_nominal_mm(inner_mm: float) -> int:
    """inner diameter (mm) → 가장 가까운 표준 호칭경.

    배관 두께 때문에 inner 가 항상 nominal 보다 작거나 같으므로 ceil 우선:
    inner 가 두 표준 호칭 사이면 큰 쪽 (안전측). 단 inner > 가장 큰 표준 이면 그 값.
    """
    if inner_mm <= 0:
        return 0
    for n in STANDARD_NOMINAL_MM:
        if inner_mm <= n + 0.5:  # 약간의 허용 오차 (78.1 ≤ 80.5)
            return n
    return STANDARD_NOMINAL_MM[-1]


def parse_sdf(path: str | Path) -> CommonNetwork:
    """SDF 파일 (XML) → CommonNetwork."""
    tree = ET.parse(str(path))
    root = tree.getroot()
    return sdf_root_to_network(root)


# SDF 의 Position attribute 단위 — emit_sdf 가 mm 로 저장 (remote30 코드 패턴).
# CommonNetwork (= KFP 호환) 는 m 단위. 즉 SDF → CommonNetwork 시 ÷ 1000 필요.
# 우리가 만든 SDF 만 mm — 외부 SDF (다이소 등) 는 m 일 가능성도 있음.
# 휴리스틱: 좌표 절대값 평균이 100 초과면 mm 추정 (도면이 100m 이상이면 mm).
def _detect_sdf_coord_unit(nodes_attr: list[tuple[float, float]]) -> float:
    """SDF Position 의 m → KFP m 변환 배율. 1.0 (이미 m) 또는 0.001 (mm → m)."""
    if not nodes_attr:
        return 1.0
    import statistics
    abs_vals = []
    for x, y in nodes_attr:
        abs_vals.append(abs(x)); abs_vals.append(abs(y))
    if not abs_vals:
        return 1.0
    mean_abs = statistics.mean(abs_vals)
    # 단위 휴리스틱: 평균 절댓값 100 초과 = mm (도면이 100m 넘는 경우 매우 드뭄)
    return 0.001 if mean_abs > 100 else 1.0


# K-Fire_Solver 의 pipe_library 와 매칭되는 표준명 — 변환 시 KFP type 표준화.
# K-solver 가 지원하는 표준 (7종):
#   KSD3507 (SPP), KSD3562 (SPPS), ASTM A135 #40,
#   Copper Type L (KS D 5301), STS Sch10S (KS D 3576),
#   CPVC (KS M 3413), K-Fire Hose
# PIPENET SDF 는 프로젝트별 단축명을 쓰는 경우가 많아 (예: "FX", "CPVC", "DP",
# "KSD 3507"), 이걸 K-solver 표준명으로 매핑. 매핑 실패 시 c-factor 휴리스틱
# (c=150→CPVC, c=130→ASTM, c=120→KSD3507) 으로 fallback.
SDF_TYPE_TO_KSOLVER = {
    # 정식 KS 명 (공백 변형)
    "KSD 3507": "KSD3507 (SPP)",
    "KSD 3562": "KSD3562 (SPPS)",
    "KSD 3576": "STS Sch10S (KS D 3576)",
    "KSD 5301": "Copper Type L (KS D 5301)",
    "KSM 3413": "CPVC (KS M 3413)",
    "ASTM A135": "ASTM A135 #40",
    # 무공백 변형
    "KSD3507": "KSD3507 (SPP)",
    "KSD3562": "KSD3562 (SPPS)",
    "KSD3576": "STS Sch10S (KS D 3576)",
    "KSD5301": "Copper Type L (KS D 5301)",
    "KSM3413": "CPVC (KS M 3413)",
    # PIPENET 프로젝트 단축명 — 17개 PIPENET SDF 샘플 분석:
    "CPVC": "CPVC (KS M 3413)",     # CPVC 단독 — KS M 3413
    "CPVC2": "CPVC (KS M 3413)",    # 권위 SLF(2. Pipenet_hand.slf) 의 CPVC schedule 명
    "FX": "KSD3507 (SPP)",          # 플렉시블 조인트 (헤드 직전 짧은 호스) — 강관 분류
    "DP": "KSD3507 (SPP)",          # Drop Pipe (헤드 직립관) — 강관
    "STS": "STS Sch10S (KS D 3576)",  # Stainless Steel
    "SPP": "KSD3507 (SPP)",
    "SPPS": "KSD3562 (SPPS)",
    "Copper": "Copper Type L (KS D 5301)",
    "Cu": "Copper Type L (KS D 5301)",
    "Hose": "K-Fire Hose",
}


def _normalize_pipe_type_name(raw_name: str, c_factor: float = 0.0) -> str:
    """KFP / K-Fire_Solver 표준 pipe type 명으로 정규화.

    1) SDF_TYPE_TO_KSOLVER 매핑 직접 lookup
    2) 매핑 실패 시 c_factor 휴리스틱 (PIPENET 일반):
       c=150 → CPVC (KS M 3413)
       c=130 → ASTM A135 #40
       c=120 → KSD3507 (SPP)  (Korean 강관 표준)
       c=100 → KSD3562 (SPPS) (Korean SPPS)
       기타 → KSD3507 (SPP)   (안전 기본)
    """
    if not raw_name:
        # c_factor 만 있으면 휴리스틱
        if c_factor >= 145: return "CPVC (KS M 3413)"
        if 125 <= c_factor < 145: return "ASTM A135 #40"
        if 95 <= c_factor < 110: return "KSD3562 (SPPS)"
        return "KSD3507 (SPP)"
    if raw_name in SDF_TYPE_TO_KSOLVER:
        return SDF_TYPE_TO_KSOLVER[raw_name]
    # 이미 K-solver 표준명 7종 중 하나면 그대로
    known = {"KSD3507 (SPP)", "KSD3562 (SPPS)", "ASTM A135 #40",
             "Copper Type L (KS D 5301)", "STS Sch10S (KS D 3576)",
             "CPVC (KS M 3413)", "K-Fire Hose"}
    if raw_name in known:
        return raw_name
    # unknown 단축명 — c_factor 휴리스틱
    if c_factor >= 145: return "CPVC (KS M 3413)"
    if 125 <= c_factor < 145: return "ASTM A135 #40"
    if 95 <= c_factor < 110: return "KSD3562 (SPPS)"
    return "KSD3507 (SPP)"


def sdf_root_to_network(root: ET.Element) -> CommonNetwork:
    """SDF XML root → CommonNetwork."""
    net = CommonNetwork(source_format="sdf")
    # Project > Network-spray (또는 직접 자식)
    net_spray = _find_child(root, "Network-spray") or root
    # 노드 — 단위 자동 감지 (우리 emit_sdf 는 mm, 외부 SDF 는 m 일 수도)
    nodes_el = _find_child(net_spray, "Nodes")
    raw_coords: list[tuple[float, float]] = []
    node_data: list = []  # (label, io, elev, x, y, pressure_bar)
    for n in _iter_children(nodes_el, "Node"):
        label = n.attrib.get("label", "")
        # ★ atmosphere outlet skip — PIPENET 이 nozzle/elastomeric outlet 을
        # <Node label="@/5"/> 로 표시. K-solver 의 노드 ID 규칙 "N<숫자>" 위반.
        # 또 미리보기에서 atmosphere 노드가 nozzle 위에 겹쳐 휘어 보이는 원인.
        if "@" in label or "/" in label or not label:
            continue
        io = n.attrib.get("io-node", "No")
        # elevation attribute 존재 여부 추적 — 우리 emit 은 항상 기록(0 도 명시)하므로
        # present 면 수리 실표고로 신뢰(0 이라도). 외부 3D SDF 는 종종 생략 → 그때만
        # Position z 를 실표고로 fallback (아래 hyd_elev_m). 표시 z 와 수리 z 분리 보장.
        elev_attr = n.attrib.get("elevation")
        has_elev = elev_attr is not None and elev_attr != ""
        elev = _ffloat(elev_attr, 0.0)
        pos = _find_child(n, "Position")
        x = _ffloat(pos.attrib.get("x") if pos is not None else 0, 0.0)
        y = _ffloat(pos.attrib.get("y") if pos is not None else 0, 0.0)
        # ★ Position z = **표시(display) 전용** 좌표(라이저 입상관 높이·헤드평면 평탄화).
        #   x,y 와 동일 coord_scale 로 처리해 평면과 비례. elevation attribute 는
        #   **수리계산용 실표고**로 분리 보존(통합 계통도는 표시 z 를 평탄화해도 head
        #   계산이 안 깨지게). Position z 없으면 elevation 을 표시 z 로 fallback.
        z_attr = pos.attrib.get("z") if pos is not None else None
        has_pos_z = z_attr is not None and z_attr != ""
        if has_pos_z:
            z_raw = _ffloat(z_attr, elev)
        else:
            z_raw = elev * 1000.0  # m → mm 일관 (x,y 가 mm 인 경우 같이 처리됨)
        cspec = _find_child(n, "Calculation-spec")
        pressure_bar = None
        water_level = None
        if cspec is not None and "pressure" in cspec.attrib:
            pressure_pa = _ffloat(cspec.attrib.get("pressure"), 101325)
            pressure_bar = pressure_pa / 100000.0
        if cspec is not None and "water-level" in cspec.attrib:
            water_level = _ffloat(cspec.attrib.get("water-level"), 0.0)
        node_data.append((label, io, elev, x, y, z_raw, has_pos_z, has_elev, pressure_bar, water_level))
        raw_coords.append((x, y))
    # ★ 단위 변환 배율 계산 — mm 추정이면 ÷ 1000
    coord_scale = _detect_sdf_coord_unit(raw_coords)
    # ★ 라벨 → K-solver 노드 ID 매핑. 순수 숫자 라벨은 N<숫자>로, 비숫자 라벨
    # (통합망의 기계실 m1/계통 n2 등) 은 충돌 없는 신규 숫자 ID 부여.
    # (이전엔 비숫자 라벨 노드를 drop 해 파이프가 댕글링 → KFP/HAS 깨짐)
    label_to_id: dict[str, str] = {}
    used_nums: set[int] = set()
    for rec in node_data:
        lbl = rec[0]
        digits = lbl.lstrip("N")
        if digits.isdigit():
            label_to_id[lbl] = lbl if lbl.startswith("N") else f"N{lbl}"
            used_nums.add(int(digits))
    next_num = (max(used_nums) + 1) if used_nums else 1
    for rec in node_data:
        lbl = rec[0]
        if lbl in label_to_id:
            continue
        while next_num in used_nums:
            next_num += 1
        label_to_id[lbl] = f"N{next_num}"
        used_nums.add(next_num)
        next_num += 1

    def _ref_id(lbl: str) -> str:
        if lbl in label_to_id:
            return label_to_id[lbl]
        return lbl if lbl.startswith("N") else f"N{lbl}"

    for label, io, elev, x, y, z_raw, has_pos_z, has_elev, pressure_bar, water_level in node_data:
        kind = "wt" if io == "Input" else "base"
        node_id = label_to_id[label]
        # 표시(display) z — Position z 있으면 x,y 와 동일 coord_scale 로(평면 비례),
        # 없으면 elevation attribute 를 표시 z 로 fallback.
        disp_z_m = (z_raw * coord_scale) if has_pos_z else elev
        # 수리 elevation — elevation attribute 가 **존재하면**(우리 emit 은 0 도 명시)
        # 그 값을 권위값으로 신뢰 → 표시 z(라이저 기둥·헤드 돌출)가 수리 표고로 새지
        # 않는다. attribute 가 아예 없는 외부 3D SDF 일 때만 Position z 로 fallback.
        hyd_elev_m = elev if (has_elev or elev != 0.0) else disp_z_m
        raw = {"sdf_label": label, "io_node": io, "display_z_m": disp_z_m}
        if water_level is not None:
            raw["water_level"] = water_level
        net.nodes[node_id] = CommonNode(
            id=node_id,
            x=x * coord_scale,
            y=y * coord_scale,
            elevation_m=hyd_elev_m,
            kind=kind,
            pressure_bar=pressure_bar,
            raw=raw,
        )

    # Links — Pipe-set > Pipe / Nozzle / Elastomeric-valve / Pump-fan
    links = _find_child(net_spray, "Links")
    # ★ 파이프 라벨 → K-solver 파이프 ID "P<숫자>" 매핑 (노드 label_to_id 와 동형).
    # K-Fire_Solver 는 int(id.lstrip('P')) 로 파싱하므로 비숫자 라벨(통합망의
    # 기계실 m1·라이저 r1 등)은 "Pm1" 이 되어 int('m1') ValueError 로 계산이
    # 멈춘다. 순수 숫자 라벨은 P<숫자> 유지, 비숫자는 충돌 없는 신규 숫자 부여.
    plabel_to_pid: dict[str, str] = {}
    used_pnums: set[int] = set()
    all_plabels: list[str] = []
    # ★ 파이프는 두 레이아웃에 존재할 수 있다:
    #   (a) PIPENET 표준 <Links><Pipe-set><Pipe/> (emit_full_sdf / 외부 PIPENET)
    #   (b) remote30 헤드검출 SDF 의 flat <Pipes><Pipe/> (Pipe-set 래퍼 없음)
    # (b) 를 놓치면 모든 노드가 고립되어 KFP/HAS 계산이 막힌다(평면도 헤드검출 경로 재현).
    # flat <Pipes> 도 pipe-group 으로 취급 — Pipe-type 자식이 없어 c-factor 휴리스틱 fallback.
    pipe_groups = list(_iter_children(links, "Pipe-set"))
    pipes_flat = _find_child(net_spray, "Pipes")
    if pipes_flat is not None:
        pipe_groups.append(pipes_flat)
    for ps in pipe_groups:
        for pipe in _iter_children(ps, "Pipe"):
            all_plabels.append(pipe.attrib.get("label", ""))
    for plbl in all_plabels:
        digits = plbl.lstrip("P")
        if digits.isdigit():
            plabel_to_pid[plbl] = plbl if plbl.startswith("P") else f"P{plbl}"
            used_pnums.add(int(digits))
    next_pnum = (max(used_pnums) + 1) if used_pnums else 1
    for plbl in all_plabels:
        if plbl in plabel_to_pid:
            continue
        while next_pnum in used_pnums:
            next_pnum += 1
        plabel_to_pid[plbl] = f"P{next_pnum}"
        used_pnums.add(next_pnum)
        next_pnum += 1

    for ps in pipe_groups:
        for pipe in _iter_children(ps, "Pipe"):
            plabel = pipe.attrib.get("label", "")
            pid = plabel_to_pid.get(
                plabel,
                plabel if plabel.startswith("P") else f"P{plabel}",
            )
            inp = pipe.attrib.get("input", "")
            out = pipe.attrib.get("output", "")
            # Pipe 가 atmosphere outlet 으로 가는 경우도 skip (nozzle 의 직접 자식 아닐 수도)
            if "@" in inp or "@" in out or "/" in inp or "/" in out:
                continue
            bore_m = _ffloat(pipe.attrib.get("bore"), 0.0)
            length = _ffloat(pipe.attrib.get("length"), 0.0)
            # PIPENET SDF 의 c-factor 는 두 곳: Pipe 의 `roughness-or-c` attribute
            # (개별 pipe 별 override), Pipe-type 의 `c-factor` (전체 기본). Pipe 우선.
            # ★ 기존엔 `c-factor` 만 봐서 Pipe override 를 놓치고 항상 type default 120
            # 으로 떨어졌음 → KFP 의 C=100 같은 값이 round-trip 에서 소실.
            c_factor = _ffloat(pipe.attrib.get("roughness-or-c"), 0.0)
            if c_factor == 0.0:
                c_factor = _ffloat(pipe.attrib.get("c-factor"), 0.0)
            # Pipe-type 검색 — 같은 Pipe-set 안에서 Name 으로 매칭
            ptype_el = _find_child(ps, "Pipe-type")
            pipe_type_label = "KSD 3507"
            if ptype_el is not None:
                ptype_name_el = _find_child(ptype_el, "Name")
                if ptype_name_el is not None and ptype_name_el.text:
                    pipe_type_label = ptype_name_el.text
                if c_factor == 0.0:
                    c_factor = _ffloat(ptype_el.attrib.get("c-factor"), 120.0)
            fittings = []
            fits_el = _find_child(pipe, "Fittings")
            for f in _iter_children(fits_el, "Fitting"):
                ftype = f.attrib.get("type", "")
                fcount = int(_ffloat(f.attrib.get("count"), 1))
                fittings.append(CommonFitting(type_id=ftype, count=fcount))
            # ★ Waypoints — PIPENET 의 휘어진 본관 중간 꺾임점들. 좌표 단위는
            # Position 과 동일하므로 coord_scale 동일 적용.
            waypoints_raw: list[tuple[float, float]] = []
            for wp in _iter_children(pipe, "Waypoints"):
                for pos in _iter_children(wp, "Position"):
                    wx = _ffloat(pos.attrib.get("x"), 0.0)
                    wy = _ffloat(pos.attrib.get("y"), 0.0)
                    waypoints_raw.append((wx, wy))
            start_id = _ref_id(inp)
            end_id = _ref_id(out)
            inner_mm = bore_m * 1000.0
            nominal_attr = pipe.attrib.get("nominal-mm")
            if nominal_attr:
                nominal_mm = int(_ffloat(nominal_attr, _nearest_nominal_mm(inner_mm)))
            else:
                nominal_mm = _nearest_nominal_mm(inner_mm)
            # ★ pipe type 명을 K-Fire_Solver 표준명으로 정규화
            #   "KSD 3507" (SDF emit) → "KSD3507 (SPP)" (K-solver pipe_library 매칭)
            # waypoints 단위 변환 — 노드 좌표와 동일 scale (Position 단위)
            waypoints_scaled = [(wx * coord_scale, wy * coord_scale)
                                  for wx, wy in waypoints_raw]
            # ★ length 는 SDF attribute 가 아니라 노드 좌표의 3D Euclidean 으로
            # 재계산 (K-solver KFP 의 length_m 규약). 사용자 제공 reference KFP 와
            # 104/104 정확히 일치. PIPENET SDF 의 length 는 waypoint 포함 routed
            # 길이라 더 길게 나옴 — K-solver 는 직선 길이 + fittings(Elbow) 로
            # 분리해 표현. 그대로 두면 length 가 ref 보다 길어 보임 (예: P10 SDF
            # 1.04 vs Ref 0.2312, 차이 0.8m).
            sn = net.nodes.get(start_id); en = net.nodes.get(end_id)
            if sn is not None and en is not None:
                length_3d = math.sqrt(
                    (sn.x - en.x) ** 2 + (sn.y - en.y) ** 2 + (sn.elevation_m - en.elevation_m) ** 2
                )
                if length_3d > 0:
                    length = length_3d
            net.pipes[pid] = CommonPipe(
                id=pid,
                start=start_id,
                end=end_id,
                length_m=length,
                diameter_inner_mm=inner_mm,
                nominal_mm=nominal_mm,
                c_factor=c_factor or 120.0,
                pipe_type_label=_normalize_pipe_type_name(pipe_type_label, c_factor),
                fittings=fittings,
                waypoints=waypoints_scaled,
                # sdf_pipe_type: 정규화 전 원명(FX/DP 등). emit_kfp 의 .slf 내경
                # lookup 이 이 원명으로 FX/DP 25.0mm 권위값을 보존하는 데 쓴다.
                raw={"sdf_label": plabel, "sdf_pipe_type": pipe_type_label},
            )

    # Nozzle — 노드를 nozzle 종류로 표시.
    # SDF 의 Nozzle 은 attribute (flow=m³/s) 또는 자식 (Flow-define/Library-item).
    # K-solver 호환 위해 head 와 nozzle 구분:
    #   - Library-item 이 "SP-HEAD" / "SPRAY" 등 head 키워드 → kind="head"
    #   - 기타 → kind="nozzle"
    for nz in net_spray.iter():
        if _strip_ns(nz.tag) != "Nozzle":
            continue
        inp_lbl = nz.attrib.get("input", "")
        node_id = _ref_id(inp_lbl)
        if node_id not in net.nodes:
            continue
        node = net.nodes[node_id]
        # 종류 추론
        node.kind = "nozzle"  # default
        flow_m3s = _ffloat(nz.attrib.get("flow"), 0.0)
        lib_item_text = ""
        for c in nz:
            tag = _strip_ns(c.tag)
            if tag == "Flow-define":
                flow_m3s = _ffloat(c.attrib.get("flow"), flow_m3s)
            elif tag == "Library-item":
                lib_item_text = (c.text or "").strip()
        if lib_item_text:
            up = lib_item_text.upper()
            # head 키워드: HEAD / SPRINKLER / SP-HEAD. 단 "SPRAY" 는 nozzle 로 (대구분).
            # 우리 emit 가 head→"SP-HEAD", nozzle→"SPRAY-NOZZLE" 로 구분.
            if ("HEAD" in up or "SPRINKLER" in up) and "SPRAY" not in up:
                node.kind = "head"
        # K 추론 — flow (m³/s) × 60000 = Q (L/min). P=1bar 가정 → K = Q.
        q_lmin = flow_m3s * 60000.0
        node.k_factor_si = q_lmin if q_lmin > 0 else 80.0  # default K=80 (SP-HEAD)
        node.raw["sdf_lib_item"] = lib_item_text or "SP-HEAD"

    # Elastomeric-valve, Pump-fan
    for vel in net_spray.iter():
        tag = _strip_ns(vel.tag)
        if tag == "Elastomeric-valve":
            inp = vel.attrib.get("input", "")
            node_id = _ref_id(inp)
            if node_id in net.nodes:
                net.nodes[node_id].kind = "valve"
                net.nodes[node_id].valve_type = vel.attrib.get("type", "alarm_valve")
        elif tag == "Pump-fan":
            inp = vel.attrib.get("input", "")
            node_id = _ref_id(inp)
            if node_id in net.nodes:
                net.nodes[node_id].kind = "pump"
                # rated-p/shutoff-p/peak-p 단위 정규화 → bar (pump_curve 내부 표준).
                #   rated-p-unit="metres" (통합망 emit_full_sdf): 양정[m] → bar.
                #   "bar" 또는 미지정(legacy KFP→SDF): 이미 bar.
                p_unit = (vel.attrib.get("rated-p-unit") or "bar").strip().lower()
                p_scale = _M_TO_BAR if p_unit == "metres" else 1.0
                rated_p_bar = _ffloat(vel.attrib.get("rated-p"), 0.0) * p_scale
                # 모든 곡선 attribute 채우기 (rated/shutoff/peak) — 압력은 bar
                net.nodes[node_id].pump_curve = {
                    "rated_q": _ffloat(vel.attrib.get("rated-q"), 0.0),
                    "rated_p": rated_p_bar,
                    "shutoff_p": _ffloat(vel.attrib.get("shutoff-p"), 0.0) * p_scale,
                    "peak_q": _ffloat(vel.attrib.get("peak-q"), 0.0),
                    "peak_p": _ffloat(vel.attrib.get("peak-p"), 0.0) * p_scale,
                    "curve_data": [],
                }
                # 펌프 토출 노드 보존 — HAS pumpBlk 의 ENodeId 는 파이프로 추론할 수
                # 없다(펌프 자체가 유일한 SRC→토출 연결이며 병렬 파이프 없음). 따라서
                # Pump-fan output 을 명시 보존해 emit_has 가 ENodeId 로 쓰게 한다.
                out_ref = _ref_id(vel.attrib.get("output", ""))
                if out_ref:
                    net.nodes[node_id].raw["pump_output"] = out_ref
                pump_name = vel.attrib.get("library-pump")
                if pump_name:
                    net.nodes[node_id].raw["pump_name"] = pump_name
                # pump 도 boundary — Calculation-spec 가능 (rated_p 는 이미 bar)
                if net.nodes[node_id].pressure_bar is None:
                    net.nodes[node_id].pressure_bar = max(rated_p_bar, 1.013)

    # Title
    title_el = _find_child(net_spray, "Title")
    if title_el is not None and title_el.text:
        net.project_meta["design_area"] = title_el.text

    # ★ 토폴로지 기반 fittings 재분배 — PIPENET SDF 는 각 junction 의 tee 를
    # 연결된 모든 pipe 에 중복 부착 (3-way 면 3개 pipe 에 각각 tee). K-solver KFP
    # 는 junction 당 한 pipe (branch 쪽) 에만 Tee(Branch). reference KFP
    # 1-1.업무시설_201동_28F 분석 결과 다음 규칙으로 33+33=66 fittings 의 80%+
    # 일치:
    #
    #   pipe.start 노드 degree=3 (분기 junction) AND pipe.end 가 head/leaf
    #     → ['Tee(Branch)']     (헤드 직전 branch 파이프 — 분기점에서 갈라짐)
    #   pipe.start degree=2 AND pipe.end 가 head/leaf
    #     → ['Elbow']            (수평관에서 헤드로 떨어지는 drop)
    #   pipe.start degree=2 AND pipe.end degree=2 AND has waypoint
    #     → ['Elbow']            (중간에 꺾이는 trunk)
    #   pipe.start degree=2 AND pipe.end degree=2 (no waypoint, but z 차이 큼)
    #     → ['Elbow']            (수직 라이저 — 직각 분기)
    #   외 모두 → []              (직진관, 또는 junction 의 through-flow 쪽)
    #
    # PIPENET 의 원본 fittings 는 cp.raw["sdf_fittings"] 로 보존 (round-trip).
    _redistribute_fittings_by_topology(net)

    return net


def _redistribute_fittings_by_topology(net: CommonNetwork) -> None:
    """파이프의 fittings 를 K-solver KFP 규약대로 토폴로지 기반 재계산.

    알고리즘 (reference KFP 1-1.업무시설_201동_28F 31 case 분석 결과):

    Step 1. PIPENET SDF 의 fittings 중 "tee" 는 junction 에 인접한 모든 파이프에
            중복 부착돼 있어 그대로 KFP 로 옮기면 N배 over-count. 따라서 일단
            모든 "tee" 를 SDF 출처 fitting 에서 제거.
    Step 2. 각 deg-3 junction 에서 perpendicularity 로 "branch" 식별. branch 의
            sdf_fittings 에 "tee" 가 있었으면 Tee(Branch) 1개 부여. 또는 branch
            의 끝이 head/leaf 면 Tee(Branch).
    Step 3. SDF "elbow" 는 designer 의 명시적 표기 → Elbow 로 유지.
            단 ① deg-3 junction 의 through-flow pipe 에 우연히 붙은 elbow 는
            제거 (junction에 elbow 가 동시에 있을 일은 거의 없음), ② 이미
            Tee(Branch) 가 있는 파이프엔 추가하지 않음.
    Step 4. deg-2 node 에서 head/leaf 로 가는 drop pipe — 아직 fitting 없으면 Elbow.
    Step 5. waypoint 만 있고 아직 fitting 없는 trunk 파이프 → Elbow.

    PIPENET 원본 fittings 는 cp.raw["sdf_fittings"] 로 보존.
    """
    from collections import defaultdict

    # 노드 degree + 인접 파이프
    node_pipes: dict[str, list[str]] = defaultdict(list)
    for pid, cp in net.pipes.items():
        node_pipes[cp.start].append(pid)
        node_pipes[cp.end].append(pid)
    deg = {n: len(v) for n, v in node_pipes.items()}

    # Step 1: 원본 SDF fittings 백업 + sdf_had_tee/elbow 마커 추출.
    # PIPENET 의 fittings 표기는 junction-중복적이라 그대로 옮기면 over-count.
    # 완전히 reset 후 토폴로지 + SDF 힌트로 재구성.
    sdf_had_tee: dict[str, bool] = {}
    sdf_had_elbow: dict[str, bool] = {}
    for pid, cp in net.pipes.items():
        cp.raw["sdf_fittings"] = [
            {"type_id": f.type_id, "count": f.count} for f in cp.fittings
        ]
        sdf_had_tee[pid] = any(f.type_id == "tee" for f in cp.fittings)
        sdf_had_elbow[pid] = any(f.type_id == "elbow" for f in cp.fittings)
        cp.fittings = []

    def _vec_from_junction(junction: str, pid: str) -> tuple[float, float]:
        cp = net.pipes[pid]
        other = cp.end if cp.start == junction else cp.start
        jn = net.nodes.get(junction); on = net.nodes.get(other)
        if jn is None or on is None:
            return (0.0, 0.0)
        return (on.x - jn.x, on.y - jn.y)

    def _angle_deg(v1, v2) -> float:
        dot = v1[0] * v2[0] + v1[1] * v2[1]
        m1 = math.sqrt(v1[0] ** 2 + v1[1] ** 2)
        m2 = math.sqrt(v2[0] ** 2 + v2[1] ** 2)
        if m1 < 1e-9 or m2 < 1e-9:
            return 0.0
        cos = max(-1.0, min(1.0, dot / (m1 * m2)))
        return math.degrees(math.acos(cos))

    # Step 2: deg-3 junction 마다 perpendicularity 로 branch 식별 → Tee(Branch)
    # (SDF "tee" 마커는 designer 의 자유로워 신뢰도 낮음, perpendicularity 가
    # 더 안정적 — 31 case 분석 결과 73/104 vs 61/104 차이.)
    for jn, pids in node_pipes.items():
        if len(pids) != 3:
            continue
        vecs = {p: _vec_from_junction(jn, p) for p in pids}
        def score(p):
            others = [o for o in pids if o != p]
            return max(_angle_deg(vecs[p], vecs[o]) for o in others)
        branch = min(pids, key=score)
        if not any(f.type_id == "tee-branch" for f in net.pipes[branch].fittings):
            net.pipes[branch].fittings.append(
                CommonFitting(type_id="tee-branch", count=1))

    # Step 3: deg-2 node 에서 head/leaf 로 가는 drop pipe (start side trunk) → Elbow
    for pid, cp in net.pipes.items():
        if cp.fittings:
            continue
        en = net.nodes.get(cp.end)
        sn = net.nodes.get(cp.start)
        end_is_leaf = en is not None and (en.kind in ("head", "nozzle") or deg.get(cp.end, 0) == 1)
        start_is_leaf = sn is not None and (sn.kind in ("head", "nozzle") or deg.get(cp.start, 0) == 1)
        if not (end_is_leaf or start_is_leaf):
            continue
        other = cp.start if end_is_leaf else cp.end
        if deg.get(other, 0) == 2:
            cp.fittings.append(CommonFitting(type_id="elbow", count=1))

    # Step 4: 파이프 내부 waypoint 가 있는데 아직 fitting 이 없으면 Elbow 부착
    for pid, cp in net.pipes.items():
        if cp.fittings:
            continue
        if len(cp.waypoints) > 0:
            cp.fittings.append(CommonFitting(type_id="elbow", count=1))


# ────────────────────────────────────────────────────────────────────────────
# SDF 직렬화 — CommonNetwork → XML
# ────────────────────────────────────────────────────────────────────────────


def _sdf_label(cn: "CommonNode", nid: str) -> str:
    """노드의 SDF label: 원본 sdf_label → 'N' prefix 제거 → 원본 id 순 fallback."""
    return cn.raw.get("sdf_label") or nid.lstrip("N") or nid


def emit_sdf_xml(net: CommonNetwork) -> str:
    """CommonNetwork → SDF XML 문자열."""
    root = ET.Element("Project", {"version": "1.8"})
    net_spray = ET.SubElement(root, "Network-spray")
    title = net.project_meta.get("design_area") or net.project_meta.get("title") or "Converted from KFP"
    ET.SubElement(net_spray, "Title").text = title

    # Nodes
    nodes_el = ET.SubElement(net_spray, "Nodes")
    for nid, cn in net.nodes.items():
        sdf_label = _sdf_label(cn, nid)
        node_attrs = {
            "label": str(sdf_label),
            "elevation": str(cn.elevation_m),
            "io-node": "Input" if cn.kind == "wt" else "No",
        }
        node_el = ET.SubElement(nodes_el, "Node", node_attrs)
        ET.SubElement(node_el, "Position", {"x": str(cn.x), "y": str(cn.y)})
        if cn.kind == "wt" and cn.pressure_bar is not None:
            # KFP bar → SDF Pa
            pa = int(round(cn.pressure_bar * 100000.0))
            ET.SubElement(node_el, "Calculation-spec", {"pressure": str(pa)})

    # Links — populated Pipe-set
    links = ET.SubElement(net_spray, "Links")
    ET.SubElement(links, "Pipe-set")  # PIPENET placeholder convention
    populated = ET.SubElement(links, "Pipe-set")
    # Pipe-type (representative) 삽입
    pipe_type_name = "KSD 3507"
    if net.pipes:
        first_pipe = next(iter(net.pipes.values()))
        pipe_type_name = first_pipe.pipe_type_label or "KSD 3507"
    ptype_el = ET.SubElement(populated, "Pipe-type", {
        "c-factor": "120", "roughness": "0.15",
    })
    ET.SubElement(ptype_el, "Name").text = pipe_type_name

    for pid, cp in net.pipes.items():
        plabel = cp.raw.get("sdf_label") or pid.lstrip("P") or pid
        in_label = (net.nodes[cp.start].raw.get("sdf_label")
                    if cp.start in net.nodes else cp.start.lstrip("N")) or cp.start
        out_label = (net.nodes[cp.end].raw.get("sdf_label")
                     if cp.end in net.nodes else cp.end.lstrip("N")) or cp.end
        pipe_attrs = {
            "label": str(plabel),
            "input": str(in_label),
            "output": str(out_label),
            "bore": str(cp.diameter_inner_mm / 1000.0),  # m
            "length": str(cp.length_m),
        }
        # ★ nominal-mm 비표준 attribute 로 호칭경 보존 — PIPENET 은 unknown
        # attribute 무시. 우리 RT 변환 시 inner ↔ nominal mismatch 회피.
        if cp.nominal_mm > 0:
            pipe_attrs["nominal-mm"] = str(cp.nominal_mm)
        pipe_el = ET.SubElement(populated, "Pipe", pipe_attrs)
        if cp.fittings:
            fits_el = ET.SubElement(pipe_el, "Fittings")
            for f in cp.fittings:
                ET.SubElement(fits_el, "Fitting", {"type": f.type_id, "count": str(f.count)})

    # Nozzle/Head 노드 → Nozzle element (둘 다 outlet)
    nozzle_idx = 1
    for nid, cn in net.nodes.items():
        if cn.kind not in ("nozzle", "head"):
            continue
        sdf_label = _sdf_label(cn, nid)
        # K [SI] L/min·bar^-0.5 → flow at default P=1bar: Q = K L/min = K/60000 m³/s
        k = cn.k_factor_si or 80.0
        flow_m3s = k / 60000.0
        nz_el = ET.SubElement(net_spray, "Nozzle", {
            "label": str(nozzle_idx),
            "input": str(sdf_label),
            "output": f"@/{nozzle_idx}",
            "flow": str(flow_m3s),
            "status": "1",
        })
        # round-trip 위해 Library-item 추가 (head/nozzle 구분 유지)
        ET.SubElement(nz_el, "Flow-define", {"flow": str(flow_m3s)})
        lib = ET.SubElement(nz_el, "Library-item")
        lib.text = "SP-HEAD" if cn.kind == "head" else "SPRAY-NOZZLE"
        nozzle_idx += 1

    # Elastomeric-valve / Pump-fan
    for nid, cn in net.nodes.items():
        if cn.kind == "valve":
            sdf_label = _sdf_label(cn, nid)
            ET.SubElement(net_spray, "Elastomeric-valve", {
                "input": str(sdf_label),
                "type": cn.valve_type or "alarm_valve",
            })
        elif cn.kind == "pump":
            sdf_label = _sdf_label(cn, nid)
            curve = cn.pump_curve or {}
            ET.SubElement(net_spray, "Pump-fan", {
                "input": str(sdf_label),
                "rated-q": str(curve.get("rated_q", 0)),
                "rated-p": str(curve.get("rated_p", 0)),
                "rated-p-unit": "bar",   # pump_curve 압력은 bar 정규화
            })

    return _pretty_xml(root)


def _pretty_xml(elem: ET.Element) -> str:
    """Pretty-print XML — indent + declaration."""
    try:
        ET.indent(elem, space="  ")
    except AttributeError:
        pass  # Python < 3.9
    return '<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(elem, encoding="unicode")


# ────────────────────────────────────────────────────────────────────────────
# 권위 SLF (2. Pipenet_hand.slf) — KFP 라이브러리의 단일 진실원천(SSOT).
# .slf 의 Schedule/Nozzle/Pump 정의를 파싱해 KFP 의 pipe/nozzle/pump_library 를
# 생성한다. 사용자가 .slf 를 모든 출력 포맷의 고정 권위 라이브러리로 지정.
# ────────────────────────────────────────────────────────────────────────────

STANDARD_SLF_FILENAME = "2. Pipenet_hand.slf"

# .slf Schedule Item-name → K-Fire_Solver 표준 7종명. K-solver 가 KFP 를 받자마자
# 매칭하도록 대응 표준이 있는 건 그 이름을 쓴다. FX/DP 는 K-solver 대응 표준이 없어
# .slf 이름을 그대로 유지(내경 25.0·조도는 .slf 권위값 보존 — 강관으로 뭉개지 않음).
_SLF_TO_KSOLVER_STD = {
    "KSD 3507": "KSD3507 (SPP)",
    "KSD 3562": "KSD3562 (SPPS)",
    "KSD 3576": "STS Sch10S (KS D 3576)",
    "CPVC2": "CPVC (KS M 3413)",
    "DP": "DP",
    "FX": "FX",
}

# Schedule 별 C-factor — .slf 는 C-factor 를 담지 않으므로(PIPENET 은 Pipe-type 에 저장)
# remote30_prototype._SCHEDULE_DEFS 와 동일 컨벤션을 따른다.
_SLF_SCHEDULE_CFACTOR = {
    "KSD 3507": 120.0, "KSD 3562": 120.0, "KSD 3576": 120.0,
    "DP": 120.0, "CPVC2": 150.0, "FX": 120.0,
}

# PIPENET K-value(m³/s·Pa^-0.5) → KFP K_SI(L/min·bar^-0.5).
#   Q[m³/s] = K_pa·√(P[Pa]) → Q[L/min] = K_pa·60000·√(P[bar]·1e5)
#            = (K_pa·60000·√1e5)·√P[bar]  →  K_si = K_pa · 60000 · √100000
_PA_K_TO_SI = 60000.0 * math.sqrt(100000.0)   # ≈ 1.8974e7


def _resolve_standard_slf() -> Path | None:
    """권위 SLF 경로 — REMOTE30_STANDARD_SLF env → 모듈 디렉토리 fallback.

    remote30_prototype.resolve_standard_slf 와 동일 규약이나, 순환 import 를 피해
    여기서 독립 해석한다.
    """
    env_val = os.environ.get("REMOTE30_STANDARD_SLF", "").strip()
    if env_val:
        cand = Path(env_val).expanduser()
        cand = cand.resolve() if cand.is_absolute() else (Path.cwd() / cand).resolve()
        if cand.is_file():
            return cand
    cand = (Path(__file__).parent / STANDARD_SLF_FILENAME).resolve()
    return cand if cand.is_file() else None


@lru_cache(maxsize=1)
def _load_slf_authoritative() -> dict:
    """권위 SLF 파싱 → {schedules, nozzles, pumps}. 파일 없으면 빈 dict.

    schedules: {item_name: {"roughness_mm": float, "sizes": {nominal_int: inner_float}}}
    nozzles:   {item_name: {"k_si": float, "max_p_pa": float, "min_p_pa": float, "desc": str}}
    pumps:     {item_name: {"curve_type": str, "flow_unit": str, "press_unit": str,
                            "points": [(flow, pressure_pa), ...], "desc": str}}
    """
    out: dict = {"schedules": {}, "nozzles": {}, "pumps": {}, "_path": None}
    slf_path = _resolve_standard_slf()
    if slf_path is None:
        return out
    out["_path"] = str(slf_path)
    try:
        root = ET.fromstring(slf_path.read_text(encoding="utf-8"))
    except (OSError, ET.ParseError):
        return out

    def _txt(parent: ET.Element | None, tag: str) -> str:
        if parent is None:
            return ""
        el = parent.find(tag)
        return (el.text or "").strip() if el is not None and el.text else ""

    sched_sec = root.find("Schedule-section")
    if sched_sec is not None:
        for sch in sched_sec.findall("Schedule"):
            name = _txt(sch, "Item-name")
            md = sch.find("Metric-definition")
            if not name or md is None:
                continue
            rough = _ffloat(md.attrib.get("roughness"), 0.0)
            sizes: dict[int, float] = {}
            for sd in md.findall("Size-definition"):
                nom = sd.attrib.get("nominal")
                inn = sd.attrib.get("internal")
                if nom is None or inn in (None, "Unset"):
                    continue
                try:
                    sizes[int(float(nom))] = float(inn)
                except (TypeError, ValueError):
                    continue
            out["schedules"][name] = {"roughness_mm": rough, "sizes": sizes}

    noz_sec = root.find("Nozzle-section")
    if noz_sec is not None:
        for nz in noz_sec.findall("Nozzle-definition"):
            name = _txt(nz, "Item-name")
            if not name:
                continue
            out["nozzles"][name] = {
                "k_si": _ffloat(nz.attrib.get("k-value"), 0.0),
                "max_p_pa": _ffloat(nz.attrib.get("maximum-pressure"), 0.0),
                "min_p_pa": _ffloat(nz.attrib.get("minimum-pressure"), 0.0),
                "desc": _txt(nz, "Description"),
            }

    pump_sec = root.find("Pump-section")
    if pump_sec is not None:
        for pd in pump_sec.findall("Pump-definition"):
            name = _txt(pd, "Item-name")
            if not name:
                continue
            points: list[tuple[float, float]] = []
            pts_el = pd.find("Set-of-pump-points")
            if pts_el is not None:
                for pp in pts_el.findall("Pump-point"):
                    points.append((_ffloat(pp.attrib.get("flow"), 0.0),
                                   _ffloat(pp.attrib.get("pressure"), 0.0)))
            out["pumps"][name] = {
                "curve_type": pd.attrib.get("curve-type", "quadratic"),
                "flow_unit": pd.attrib.get("flowrate-unit", "l-min"),
                "press_unit": pd.attrib.get("pressure-unit", "metres"),
                "points": points,
                "desc": _txt(pd, "Description"),
            }
    return out


def _slf_inner_d_mm(nominal_mm: int, slf_type_name: str) -> float | None:
    """권위 SLF schedule 에서 (호칭 → 내경 mm). 매칭 실패 시 None.

    slf_type_name 은 .slf 원명("FX","DP","KSD 3507"...) 또는 K-solver 정규화명
    ("KSD3507 (SPP)"...) 둘 다 허용 — 역매핑으로 .slf 원명을 찾는다.
    """
    slf = _load_slf_authoritative()
    scheds = slf.get("schedules") or {}
    if not scheds:
        return None
    # 1) .slf 원명 직접
    if slf_type_name in scheds:
        return scheds[slf_type_name]["sizes"].get(int(nominal_mm))
    # 2) K-solver 정규화명 → .slf 원명 역매핑
    for slf_name, ks_name in _SLF_TO_KSOLVER_STD.items():
        if ks_name == slf_type_name and slf_name in scheds:
            v = scheds[slf_name]["sizes"].get(int(nominal_mm))
            if v is not None:
                return v
    return None


def _slf_roughness_mm(slf_type_name: str) -> float | None:
    """권위 SLF schedule 의 조도(mm). 매칭 실패 시 None."""
    slf = _load_slf_authoritative()
    scheds = slf.get("schedules") or {}
    if slf_type_name in scheds:
        return scheds[slf_type_name].get("roughness_mm")
    for slf_name, ks_name in _SLF_TO_KSOLVER_STD.items():
        if ks_name == slf_type_name and slf_name in scheds:
            return scheds[slf_name].get("roughness_mm")
    return None


def _slf_pipe_library_for_kfp() -> dict | None:
    """권위 SLF schedules → KFP pipe_library dict. .slf 없으면 None."""
    slf = _load_slf_authoritative()
    scheds = slf.get("schedules") or {}
    if not scheds:
        return None
    pipe_types: list[dict] = []
    for slf_name, sdef in scheds.items():
        std = _SLF_TO_KSOLVER_STD.get(slf_name, slf_name)
        cf = _SLF_SCHEDULE_CFACTOR.get(slf_name, 120.0)
        rough = sdef.get("roughness_mm", 0.0)
        for nom, inner in sorted(sdef.get("sizes", {}).items()):
            pipe_types.append({
                "standard": std,
                "nominal_mm": int(nom),
                "inner_d_mm": float(inner),
                "roughness_mm": float(rough),
                "C": float(cf),
            })
    return {
        "version": "slf-authoritative",
        "description": f"Generated from {Path(slf['_path']).name}" if slf.get("_path") else "Generated from SLF",
        "pipe_types": pipe_types,
    }


def _slf_nozzle_library_for_kfp() -> dict | None:
    """권위 SLF nozzles → KFP nozzle_library (head/nozzle 카테고리). .slf 없으면 None."""
    slf = _load_slf_authoritative()
    nozzles = slf.get("nozzles") or {}
    if not nozzles:
        return None
    head_items = []
    for name, ndef in nozzles.items():
        k_si = ndef.get("k_si", 0.0) * _PA_K_TO_SI
        head_items.append({
            "id": f"HEAD_{name.replace(' ', '_').replace('/', '_').upper()}",
            "display_name": name,
            "K_SI": round(k_si, 4),
            "K_US": round(k_si / 1.4159, 4),
            "min_p_bar": round(ndef.get("min_p_pa", 0.0) / 1e5, 5),
            "max_p_bar": round(ndef.get("max_p_pa", 0.0) / 1e5, 5),
            "is_system_protected": True,
        })
    return {
        "schema_version": "2.0",
        "categories": [{
            "category_id": "head",
            "display_name": "Head",
            "is_system_protected": True,
            "type_id": "head",
            "items": head_items,
        }],
    }


def _slf_pump_library_for_kfp() -> list[dict]:
    """권위 SLF pumps → KFP pump_library pumps 리스트. .slf 없으면 빈 리스트.

    .slf Pump-point pressure 는 Pa(레퍼런스 실측). KFP pump_library 압력은 bar.
    points 에서 rated(중간점)/shutoff(flow0)/peak(max flow) 를 도출한다.
    """
    slf = _load_slf_authoritative()
    pumps = slf.get("pumps") or {}
    out: list[dict] = []
    for name, pdef in pumps.items():
        pts = pdef.get("points") or []
        if not pts:
            continue
        pts_sorted = sorted(pts, key=lambda fp: fp[0])
        flow0 = pts_sorted[0]
        flowmax = pts_sorted[-1]
        mid = pts_sorted[len(pts_sorted) // 2]
        out.append({
            "id": f"SLF_{name.replace(' ', '_').upper()}",
            "name": name,
            "description": pdef.get("desc", ""),
            "curve_type": pdef.get("curve_type", "quadratic"),
            "rated_q": round(mid[0], 6),
            "rated_p": round(mid[1] / 1e5, 5),
            "shutoff_p": round(flow0[1] / 1e5, 5),
            "peak_q": round(flowmax[0], 6),
            "peak_p": round(flowmax[1] / 1e5, 5),
            "curve_points_bar": [[round(f, 6), round(p / 1e5, 5)] for f, p in pts_sorted],
        })
    return out


# ────────────────────────────────────────────────────────────────────────────
# K-Fire_Solver 표준 라이브러리 — _internal/{nozzle,pipe,fittings}_library*.json
# 내용을 모듈 상수로 embed. emit_kfp 가 출력에 그대로 포함해 K-Fire_Solver 가
# 매칭 가능하게 함. SDF → KFP 변환 시 default 빈 라이브러리 대신 이걸 사용.
# fittings 만 K-solver(NFPA L/D) 유지 — .slf Fitting-section 이 비어 매칭원천이 없음.
# ────────────────────────────────────────────────────────────────────────────

KSOLVER_NOZZLE_LIBRARY = {
    "nozzles": [
        {"name": "Hose Nozzle 13mm", "P_bar": 0.0, "Q_lpm": 130.0, "K_val": 99.7, "min_p": 1.7},
        {"name": "Hose Nozzle 19mm", "P_bar": 0.0, "Q_lpm": 400.0, "K_val": 214.0, "min_p": 3.5},
        {"name": "Water Spray Type-A", "P_bar": 3.5, "Q_lpm": 120.0, "K_val": 64.1, "min_p": 3.5},
        {"name": "Water Spray Type-B", "P_bar": 3.5, "Q_lpm": 60.0, "K_val": 32.1, "min_p": 3.5},
        {"name": "Water Spray Type-C", "P_bar": 3.5, "Q_lpm": 90.0, "K_val": 48.1, "min_p": 3.5},
    ],
}

@lru_cache(maxsize=1)
def _try_load_ksolver_libraries() -> dict:
    """K-Fire_Solver/_internal/*.json 가 있으면 그걸 우선 사용. 없으면 embedded.

    파이프마다 호출돼 JSON 을 매번 재파싱하던 핫스팟 — 프로세스 생애 1회만 읽도록
    캐시. 반환 dict 는 read-only 로만 사용한다(호출처 전부 .get 조회).
    """
    out = {
        "nozzle": KSOLVER_NOZZLE_LIBRARY,
        "pipe": None,
        "fittings": None,
    }
    ksolver_dir = Path(__file__).parent / "K-Fire_Solver" / "_internal"
    if ksolver_dir.exists():
        for key, fname in [
            ("nozzle", "nozzle_library.json"),
            ("pipe", "pipe_library.json"),
            ("fittings", "fittings_library_from_nfpa_ld.json"),
        ]:
            p = ksolver_dir / fname
            if p.exists():
                try:
                    out[key] = json.loads(p.read_text(encoding="utf-8"))
                except Exception:
                    pass
    # fallback embedded
    if out["pipe"] is None:
        out["pipe"] = {
            "version": "1.1",
            "description": "Embedded fallback (K-Fire_Solver 폴더 없음)",
            "pipe_types": [],  # 최소화 — 실제 작업 시 K-Fire_Solver 폴더가 있어야
        }
    if out["fittings"] is None:
        out["fittings"] = {"fittings": []}
    return out


_FITTINGS_V3_CACHE: list | None = None


def _fittings_library_v3() -> list:
    """4.0-NFPA13-EQL 피팅 카탈로그 — 번들 kfp_fittings_library_v3.json 로드(1회 캐시).

    K-Fire_Solver 4.0 은 KFP 로드 시 fittings_library_v3(NFPA13 등가길이 표, DN별
    equivalent_lengths)를 요구한다. 2.2 의 L/D 기반 fittings_library 는 4.0 이 안 읽음.
    카탈로그는 ewqwqe.kfp(4.0 빈 템플릿)에서 추출한 권위 데이터.
    """
    global _FITTINGS_V3_CACHE
    if _FITTINGS_V3_CACHE is None:
        p = Path(__file__).parent / "kfp_fittings_library_v3.json"
        try:
            _FITTINGS_V3_CACHE = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            _FITTINGS_V3_CACHE = []
    return copy.deepcopy(_FITTINGS_V3_CACHE)


def _lookup_inner_d_mm(nominal_mm: int, standard: str) -> float:
    """K-solver pipe_library 에서 standard + nominal → inner_d_mm 매핑.
    매칭 실패 시 nominal_mm 그대로 반환."""
    libs = _try_load_ksolver_libraries()
    pipe_types = libs["pipe"].get("pipe_types", [])
    for pt in pipe_types:
        if pt.get("nominal_mm") == nominal_mm and pt.get("standard") == standard:
            return float(pt.get("inner_d_mm", nominal_mm))
    # fallback: 같은 nominal 가진 첫 항목
    for pt in pipe_types:
        if pt.get("nominal_mm") == nominal_mm:
            return float(pt.get("inner_d_mm", nominal_mm))
    return float(nominal_mm)


_FITTING_NAME_ALIASES = {
    # 단순 이름 → K-solver 표준명
    "elbow": "90 deg Standard Elbow",
    "tee": "Tee / Cross (90 deg branch flow)",
    "tee(branch)": "Tee / Cross (90 deg branch flow)",
    "tee branch": "Tee / Cross (90 deg branch flow)",
    "tee_branch": "Tee / Cross (90 deg branch flow)",
    "tee(run)": "Tee / Cross (90 deg branch flow)",  # run flow — L/D 다르지만 일단 같은 카탈로그
    "cross": "Tee / Cross (90 deg branch flow)",
    "elbow 45": "45 deg Elbow",
    "elbow 90": "90 deg Standard Elbow",
    "elbow long": "90 deg Long Radius Elbow",
    "gate": "Gate Valve",
    "butterfly": "Butterfly Valve",
    "check": "Swing Check Valve",
    "alarm": "Alarm Valve",
    "dry": "Dry Pipe Valve",
    "preaction": "Preaction Valve",
    "ball": "Ball Valve",
    "angle": "Angle Valve",
    "glove": "Glove Valve",
    "globe": "Glove Valve",
}


def _lookup_fitting_lod(name: str) -> float | None:
    """K-solver fittings_library 에서 name → L_over_D 매핑.

    매칭 우선순위:
    1. alias dict (단순 "Elbow" → "90 deg Standard Elbow")
    2. 정확한 name match
    3. substring match
    """
    libs = _try_load_ksolver_libraries()
    fits = libs["fittings"].get("fittings", [])
    if not fits:
        return None
    target = (name or "").strip().lower()
    # 1. alias
    aliased = _FITTING_NAME_ALIASES.get(target)
    if aliased:
        for f in fits:
            if f.get("name") == aliased:
                return float(f.get("L_over_D", 0))
    # 2. 정확 name (case-insensitive)
    for f in fits:
        if f.get("name", "").lower() == target:
            return float(f.get("L_over_D", 0))
    # 3. substring
    for f in fits:
        if target in f.get("name", "").lower() or target == f.get("type", "").lower():
            return float(f.get("L_over_D", 0))
    return None


def _add_boundary_calc_spec(sdf_path: Path, nodes: list[dict],
                              slf_filename: str | None = None,
                              pumps: list[dict] | None = None) -> None:
    """SDF 후처리 — PIPENET 의 핵심 reject 두 가지 동시 해결.

    (1) io-node="Input" 노드에 <Calculation-spec pressure=Pa/> 추가:
        "No analysis pressure specification" 에러 직접 해결.
    (2) <User-lib file="..."> 의 file attribute 에 SLF 파일명 명시:
        PIPENET 이 이 경로의 SLF 를 열어 nozzle K-factor / pipe-size 매핑.
        - 다이소 reference: file="Y:\\...절대경로...\\1-1.다이소....slf"
        - 우리: 단순 파일명 ("tree25.slf") — 사용자가 ZIP 풀면 같은 폴더에 있어
          PIPENET 이 자동 인식. "Nozzle K Factor must be given" / "Pipe bore
          must be given" 에러는 SLF 경로 매칭 실패가 직접 원인.
    """
    import xml.etree.ElementTree as _ET
    tree = _ET.parse(str(sdf_path))
    root = tree.getroot()
    # (1) Calculation-spec — pressure + water-level (wt 고가수조)
    pressure_by_label = {
        str(n["label"]): int(n["pressure_pa"])
        for n in nodes
        if n.get("pressure_pa") is not None and n.get("io_node") == "Input"
    }
    water_level_by_label = {
        str(n["label"]): float(n["water_level"])
        for n in nodes
        if n.get("water_level") is not None
    }
    for node_el in root.iter("Node"):
        label = node_el.get("label", "")
        if label not in pressure_by_label:
            continue
        has_spec = any(c.tag.endswith("Calculation-spec") for c in node_el)
        if has_spec:
            # 기존 spec 에 water-level attribute 추가 (round-trip 시 누락 방지)
            for c in node_el:
                if c.tag.endswith("Calculation-spec") and label in water_level_by_label:
                    c.set("water-level", str(water_level_by_label[label]))
                    break
            continue
        attrs = {"pressure": str(pressure_by_label[label])}
        if label in water_level_by_label:
            attrs["water-level"] = str(water_level_by_label[label])
        _ET.SubElement(node_el, "Calculation-spec", attrs)
    # (2) User-lib file
    if slf_filename:
        for ul in root.iter("User-lib"):
            ul.set("file", slf_filename)
    # (★) Position 복원 — remote30_prototype.emit_sdf 의 _xform 이 좌표를
    # (-cx, -cy) shift + 3000/longest scale 로 재정규화해 PIPENET 캔버스에 fit
    # 시켰는데, 그 결과 원본 KFP 좌표가 손실됨 (예: tree25 의 63 노드가 31 unique
    # 좌표로 붕괴 → K-Fire_Solver 가 "Error 223 not enough nodes"). 사용자 지시:
    # "솔버의 x축 = PIPENET 의 수평, 솔버의 y축 = PIPENET 의 수직" — 즉 좌표를
    # 그대로 통과시켜야 함. 여기서 nodes (tables.nodes) 에 있던 원본 mm 좌표를
    # Position 에 덮어 쓴다.
    # z 도 같이 처리 — KFP nodes 가 [x, y, z] 라 z = elevation, 단위 mm.
    label_to_xyz = {
        str(n["label"]): (
            float(n.get("x", 0.0)),
            float(n.get("y", 0.0)),
            float(n.get("elevation", 0.0)) * 1000.0,
        )
        for n in nodes
    }
    for node_el in root.iter("Node"):
        lbl = node_el.get("label", "")
        if lbl not in label_to_xyz:
            continue
        x_mm, y_mm, z_mm = label_to_xyz[lbl]
        for pos in node_el:
            if pos.tag.endswith("Position"):
                pos.set("x", str(x_mm))
                pos.set("y", str(y_mm))
                pos.set("z", str(z_mm))
                break

    # (3a) Elastomeric-valve emit — emit_sdf 가 valve 노드를 base 로 떨어뜨려
    # round-trip 시 valve 정보 소실. 후처리에서 Network-spray 안에 추가.
    # valve 정보는 nodes dict 안의 "valve_type" 필드에 담겨 옴.
    valve_nodes = [n for n in nodes if n.get("valve_type")]
    if valve_nodes:
        for ns in root.iter():
            if ns.tag.endswith("Network-spray"):
                for vn in valve_nodes:
                    vtype = vn["valve_type"]
                    _ET.SubElement(ns, "Elastomeric-valve", {
                        "input": str(vn["label"]),
                        "output": str(vn["label"]),  # in-place valve (same node)
                        "type": str(vtype),
                    })
                break

    # (3b) Pump-fan emit — emit_sdf 가 안 만들기 때문에 후처리로 추가.
    # SDF→KFP round-trip 시 pump 정보 보존을 위해 필수.
    if pumps:
        # Pump-fan element 는 Network-spray 안에. 첫 Network-spray 찾기.
        for ns in root.iter():
            if ns.tag.endswith("Network-spray"):
                for pump in pumps:
                    pf = _ET.SubElement(ns, "Pump-fan", {
                        "input": str(pump["label"]),
                        "rated-q": str(pump.get("rated_q", 0)),
                        "rated-p": str(pump.get("rated_p", 0)),
                        "shutoff-p": str(pump.get("shutoff_p", 0)),
                        "peak-q": str(pump.get("peak_q", 0)),
                        "peak-p": str(pump.get("peak_p", 0)),
                        "rated-p-unit": "bar",   # pump_curve 압력은 bar 정규화
                    })
                    if pump.get("name"):
                        pf.set("library-pump", str(pump["name"]))
                break
    # DOCTYPE 보존 — ET.write 는 <!DOCTYPE Project SYSTEM "spray.dtd"> 를 떨어뜨려
    # 다른 PC 의 PIPENET 에서 파일 열기/연산이 거부될 수 있다(레퍼런스 SDF 헤더 정합).
    _body = _ET.tostring(root, encoding="unicode")
    Path(sdf_path).write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE Project SYSTEM "spray.dtd">\n' + _body,
        encoding="utf-8",
    )


def network_to_pipe_tables(net: CommonNetwork):
    """CommonNetwork → remote30_prototype.PipeTables 변환.

    remote30_prototype.emit_sdf 가 PIPENET native SDF (template-based,
    6 schedule + SLF 동봉 + Graphics 메타) 를 emit 하므로, 우리 KFP 파싱
    결과를 그 함수가 받을 수 있는 형태로 변환한다. 단순 XML 직렬화 (`emit_sdf_xml`)
    보다 PIPENET 호환성이 훨씬 좋음 (계산 + 아이소 다 통과).
    """
    from remote30_prototype import PipeTables  # 지연 import — 순환 회피
    tables = PipeTables()

    # 노드 라벨링 — KFP "N5" → SDF "5" (prefix 제거).
    # io_node="Input" 후보: wt + pump (둘 다 system source/boundary).
    node_label_map: dict[str, str] = {}
    for nid, cn in net.nodes.items():
        sdf_label = cn.raw.get("sdf_label") or (nid[1:] if nid.startswith("N") else nid)
        node_label_map[nid] = str(sdf_label)
        io_node = "Input" if cn.kind in ("wt", "pump") else "No"
        node_dict = {
            "label": str(sdf_label),
            "elevation": cn.elevation_m,
            "io_node": io_node,
            "x": int(round(cn.x * 1000)),   # SDF emit 은 mm 단위 좌표 가정 (remote30 코드 패턴)
            "y": int(round(cn.y * 1000)),
        }
        # ★ Calculation-spec — boundary 노드의 pressure_pa.
        # 이게 없으면 PIPENET "No analysis pressure specification" 에러.
        if io_node == "Input" and cn.pressure_bar is not None:
            node_dict["pressure_pa"] = int(round(cn.pressure_bar * 100000))
        elif io_node == "Input":
            node_dict["pressure_pa"] = 101325  # 대기압 fallback
        # ★ valve node — valve_type 보존 (KFP↔SDF round-trip 시 Elastomeric-valve
        # element 로 emit 되어야 valve 정체성이 보존됨).
        if cn.kind == "valve" and cn.valve_type:
            node_dict["valve_type"] = str(cn.valve_type)
        # ★ wt node — water_level 보존 (K-solver 고가수조 driving head 의 핵심).
        # KFP→SDF→KFP round-trip 시 water_level 이 0 으로 떨어지면 K-solver 가
        # 펌프 없는 시스템에서 정수두 0 이 되어 flow 계산 불능 ("계산 안됨").
        # SDF 의 Calculation-spec 에 custom attribute water-level 로 동봉.
        if cn.kind == "wt":
            wl = cn.raw.get("water_level")
            if wl is not None:
                node_dict["water_level"] = float(wl)
        tables.nodes.append(node_dict)

    # 파이프 — KFP "P7" → SDF "7" (prefix 제거).
    for pid, cp in net.pipes.items():
        sdf_plabel = cp.raw.get("sdf_label") or (pid[1:] if pid.startswith("P") else pid)
        tables.pipes.append({
            "label": str(sdf_plabel),
            "in": node_label_map.get(cp.start, cp.start.lstrip("N")),
            "out": node_label_map.get(cp.end, cp.end.lstrip("N")),
            "type": cp.pipe_type_label or "KSD 3507",
            "dia": cp.nominal_mm or _nearest_nominal_mm(cp.diameter_inner_mm),
            "length": max(round(cp.length_m, 3), 0.001),
            "elev": 0.0,
            "c": str(int(round(cp.c_factor))),
            "status": "Normal",
            "group": "Unset",
        })
        # 파이프 안 fitting → tables.fittings (PIPENET native 는 pipe 자식)
        for f in cp.fittings:
            tables.fittings.append({
                "pipe": str(sdf_plabel),
                "in": node_label_map.get(cp.start, cp.start.lstrip("N")),
                "out": node_label_map.get(cp.end, cp.end.lstrip("N")),
                "type": f.type_id or "elbow",
                "count": f.count,
            })

    # ★ 노즐 — `head` 와 `nozzle` 둘 다 outlet 처리.
    # PIPENET 의 "Network must have outlets or nozzles" 에러 직접 원인:
    # 우리가 nozzle 만 처리해서 head 노드 (k_factor_si 보유) 누락.
    nozzle_idx = 1
    for nid, cn in net.nodes.items():
        if cn.kind not in ("nozzle", "head"):
            continue
        k = cn.k_factor_si or 80.0
        flow_lmin = k  # K [L/min·bar^-0.5] × √1 = L/min @ P=1bar
        # head 와 nozzle 의 SDF library 구분 — head 는 표준 SP-HEAD, nozzle 은
        # 분무 노즐 (open SP). PIPENET 의 SLF library 매칭 위해 다른 lib id.
        lib_name = "SP-HEAD" if cn.kind == "head" else "OPEN-SP"
        tables.nozzles.append({
            "label": str(nozzle_idx),
            "in": node_label_map[nid],
            "out": f"@/{nozzle_idx}",
            "status": "1",
            "lib": lib_name,
            "flow_m3s": flow_lmin / 60000.0,
            "flow_lmin": flow_lmin,
        })
        nozzle_idx += 1

    # Meta
    tables.meta = [
        ("Source format", "KFP (K-solver)"),
        ("Conversion target", "PIPENET SDF native"),
        ("Project", net.project_meta.get("design_area", "")),
        ("Node count", str(len(tables.nodes))),
        ("Pipe count", str(len(tables.pipes))),
        ("Nozzle count", str(len(tables.nozzles))),
    ]
    return tables


def convert_kfp_to_sdf(
    kfp_path: str | Path,
    sdf_path: str | Path | None = None,
    *,
    use_pipenet_native: bool = True,
    project_title: str = "Converted from KFP",
    slf_filename: str | None = None,
) -> str:
    """KFP → SDF.

    use_pipenet_native=True (기본): remote30_prototype.emit_sdf 활용.
    template-based 직렬화로 PIPENET 의 Attributes / Libraries / Graphics 메타
    (아이소매트릭 표시 등) 보존. SLF 동봉 + 6 schedule embed + 빈 Pipe-set
    placeholder 컨벤션 자동 적용. PIPENET 에서 계산 + 시각화 모두 통과.

    use_pipenet_native=False: 단순 XML 직렬화 (`emit_sdf_xml`). 의미 보존은
    같지만 PIPENET native 메타 누락. 디버그/round-trip 검증 용.
    """
    net = parse_kfp(kfp_path)
    if use_pipenet_native:
        # remote30_prototype.emit_sdf 활용 — 임시 파일 통해 PIPENET native 형식 생성,
        # 그 후 Calculation-spec / Pump-fan / Elastomeric-valve 후처리 추가
        # (emit_full_sdf 의 후처리 패턴 차용 — boundary pressure 필수).
        import tempfile
        from remote30_prototype import emit_sdf as _emit_sdf
        tables = network_to_pipe_tables(net)
        # pump 노드 정보 추출 (Pump-fan emit 용)
        pumps_for_post = [
            {
                "label": str(cn.raw.get("sdf_label") or nid.lstrip("N")),
                "name": cn.raw.get("pump_name") or "Default Pump",
                **(cn.pump_curve or {}),
            }
            for nid, cn in net.nodes.items() if cn.kind == "pump"
        ]
        if sdf_path is None:
            with tempfile.NamedTemporaryFile(suffix=".sdf", delete=False) as _tmp:
                tmp_path = Path(_tmp.name)
            try:
                _emit_sdf(tables, tmp_path, project_title=project_title)
                _add_boundary_calc_spec(tmp_path, tables.nodes,
                                          slf_filename=slf_filename,
                                          pumps=pumps_for_post)
                xml_str = tmp_path.read_text(encoding="utf-8")
            finally:
                try: tmp_path.unlink()
                except OSError: pass
        else:
            _emit_sdf(tables, Path(sdf_path), project_title=project_title)
            _add_boundary_calc_spec(Path(sdf_path), tables.nodes,
                                      slf_filename=slf_filename,
                                      pumps=pumps_for_post)
            xml_str = Path(sdf_path).read_text(encoding="utf-8")
        return xml_str
    # Fallback — 단순 XML
    xml_str = emit_sdf_xml(net)
    if sdf_path is not None:
        Path(sdf_path).write_text(xml_str, encoding="utf-8")
    return xml_str


def convert_kfp_to_sdf_with_slf(kfp_path: str | Path, out_dir: str | Path,
                                  project_title: str = "Converted from KFP") -> dict:
    """KFP → SDF + SLF 동봉. 결과 디렉토리에 .sdf 와 .slf 같이 저장.

    PIPENET 은 schedule 라이브러리 (.slf) 가 .sdf 와 같은 폴더에 있어야
    내경(Internal) Unset 이슈 없이 정확히 동작. REMOTE30_STANDARD_SLF env
    또는 모듈 디렉토리 fallback 에서 표준 SLF 찾아 복사.
    """
    import shutil
    from remote30_prototype import resolve_standard_slf  # 표준 SLF 경로 해석
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(kfp_path).stem
    sdf_out = out_dir / f"{stem}.sdf"
    slf_out = out_dir / f"{stem}.slf"
    # SDF 생성
    convert_kfp_to_sdf(kfp_path, sdf_out, use_pipenet_native=True,
                         project_title=project_title)
    # 표준 SLF 복사
    try:
        slf_src = resolve_standard_slf()
        shutil.copy(slf_src, slf_out)
        slf_status = f"copied from {slf_src}"
    except Exception as exc:
        slf_status = f"SLF 동봉 실패: {exc}"
    return {"sdf": str(sdf_out), "slf": str(slf_out), "slf_status": slf_status}


def convert_sdf_to_kfp(sdf_path: str | Path, kfp_path: str | Path | None = None,
                       *, coord_scale: float = 1.0,
                       display_geometry: bool = False,
                       simplify: bool = True) -> dict:
    """SDF → KFP dict. kfp_path 주어지면 파일 저장. coord_scale 은 표시좌표 배율.

    display_geometry — 통합망 전용(미리보기 비례 스키매틱 좌표). emit_kfp 참조.
    simplify — 일직선 통과 base 노드 병합으로 출력 노드 수 절감(무손실). 기본 True.
    """
    net = parse_sdf(sdf_path)
    if simplify:
        simplify_passthrough_nodes(net)
    out = emit_kfp(net, kfp_path, coord_scale=coord_scale,
                   display_geometry=display_geometry)
    return out


def parse_for_preview(path: str | Path) -> dict:
    """KFP 또는 SDF 파일 → 미리보기 데이터.

    포맷 자동 감지: 첫 바이트가 '{' 면 JSON (KFP), '<' 면 XML (SDF).
    반환 구조 (캔버스 그릴 수 있는 형태):
      {
        "format": "kfp" | "sdf",
        "nodes": [{"id":..., "x":..., "y":..., "z":..., "kind":...}],
        "edges": [{"id":..., "a":..., "b":..., "dn":..., "length":...}],
        "meta": {"node_count":..., "pipe_count":..., "nozzle_count":..., ...},
        "bbox": [xmin, ymin, xmax, ymax],
      }
    """
    p = Path(path)
    head = p.read_bytes()[:8].lstrip(b"\xef\xbb\xbf").lstrip()
    fmt = "kfp" if head[:1] in (b"{", b"[") else "sdf"
    net = parse_kfp(p) if fmt == "kfp" else parse_sdf(p)
    nodes_view = []
    xs = []; ys = []
    for nid, cn in net.nodes.items():
        nodes_view.append({
            "id": nid, "x": cn.x, "y": cn.y, "z": cn.elevation_m,
            "kind": cn.kind,
            "k_factor": cn.k_factor_si,
            "valve_type": cn.valve_type,
        })
        xs.append(cn.x); ys.append(cn.y)
    # waypoints 좌표도 bbox 포함 — 휘어진 본관이 노드 범위를 넘어가는 경우 화면 잘림 방지
    for cp in net.pipes.values():
        for wx, wy in cp.waypoints:
            xs.append(wx); ys.append(wy)
    edges_view = []
    for pid, cp in net.pipes.items():
        edges_view.append({
            "id": pid, "a": cp.start, "b": cp.end,
            "dn": cp.nominal_mm, "length": cp.length_m,
            "c_factor": cp.c_factor, "type": cp.pipe_type_label,
            "fitting_count": sum(f.count for f in cp.fittings),
            # ★ waypoints — start → wp1 → ... → end 폴리라인 경로 (m). 미리보기
            # 캔버스가 직선 대신 path 따라 그려야 휘어진 본관이 정확히 보임.
            "waypoints": cp.waypoints,
        })
    # Robust bbox — 0.5%/99.5% percentile 로 outlier 노드 (atmosphere 잔재, 표제란 등)
    # 영향 제거. raw min/max 만 쓰면 노드 하나가 멀리 떨어져도 bbox 폭주 → 캔버스
    # fit 시 메인 도면이 작은 점으로 보이거나 휘어 보이는 원인.
    if xs and ys:
        xs_sorted = sorted(xs); ys_sorted = sorted(ys)
        n = len(xs_sorted)
        lo = max(int(n * 0.005), 0)
        hi = min(int(n * 0.995), n - 1)
        bbox = [xs_sorted[lo], ys_sorted[lo], xs_sorted[hi], ys_sorted[hi]]
        # 1% margin
        w = bbox[2] - bbox[0]; h = bbox[3] - bbox[1]
        mx = max(w * 0.05, 0.1); my = max(h * 0.05, 0.1)
        bbox = [bbox[0] - mx, bbox[1] - my, bbox[2] + mx, bbox[3] + my]
    else:
        bbox = [0.0, 0.0, 1.0, 1.0]
    kind_counts: dict[str, int] = {}
    for cn in net.nodes.values():
        kind_counts[cn.kind] = kind_counts.get(cn.kind, 0) + 1
    return {
        "format": fmt,
        "nodes": nodes_view,
        "edges": edges_view,
        "bbox": bbox,
        "meta": {
            "node_count": len(nodes_view),
            "pipe_count": len(edges_view),
            "kind_counts": kind_counts,
            "nozzle_library_count": len(net.nozzle_library),
            "fitting_library_count": len(net.fitting_library),
            "pump_library_count": len(net.pump_library),
            "design_settings": net.design_settings,
            "project_title": net.project_meta.get("design_area", ""),
        },
    }


# ────────────────────────────────────────────────────────────────────────────
# 자가 검증 — round-trip 손실 측정
# ────────────────────────────────────────────────────────────────────────────


def round_trip_check(kfp_path: str | Path) -> dict:
    """KFP → SDF → KFP round-trip, 손실 측정.

    반환: {
      "node_count_orig": int, "node_count_rt": int,
      "pipe_count_orig": int, "pipe_count_rt": int,
      "node_kind_diff": {...},
      "pipe_length_rmse": float (m),
      "pipe_dn_match_pct": float (0~100),
    }
    """
    orig = parse_kfp(kfp_path)
    xml = emit_sdf_xml(orig)
    # SDF parse 다시
    rt_net = sdf_root_to_network(ET.fromstring(xml))

    orig_kinds = sorted([cn.kind for cn in orig.nodes.values()])
    rt_kinds = sorted([cn.kind for cn in rt_net.nodes.values()])
    from collections import Counter
    diff = {}
    for k in set(orig_kinds) | set(rt_kinds):
        diff[k] = (Counter(orig_kinds).get(k, 0), Counter(rt_kinds).get(k, 0))

    # Pipe length / dn 비교 (id 매칭)
    rt_by_id = {p.id: p for p in rt_net.pipes.values()}
    length_diffs = []
    dn_matches = 0
    dn_total = 0
    for pid, op in orig.pipes.items():
        rp = rt_by_id.get(pid)
        if rp is None:
            continue
        length_diffs.append((op.length_m - rp.length_m) ** 2)
        dn_total += 1
        if op.nominal_mm == rp.nominal_mm:
            dn_matches += 1
    length_rmse = math.sqrt(sum(length_diffs) / max(1, len(length_diffs)))

    return {
        "node_count_orig": len(orig.nodes),
        "node_count_rt": len(rt_net.nodes),
        "pipe_count_orig": len(orig.pipes),
        "pipe_count_rt": len(rt_net.pipes),
        "node_kind_diff": diff,
        "pipe_length_rmse_m": round(length_rmse, 4),
        "pipe_dn_match_pct": round(100 * dn_matches / max(1, dn_total), 2),
    }

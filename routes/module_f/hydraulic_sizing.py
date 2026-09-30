"""Explicit adapter from confirmed merge tables to the independent solver."""
from __future__ import annotations

import hashlib
import json
import math
from typing import Any

from src.pipenet_converter.hydraulics.models import Network, Node, Nozzle, Pipe, Settings, Size
from src.pipenet_converter.graph.fitting_review import fitting_review


def fingerprint(tables: Any) -> str:
    """Include physics, connectivity, evidence and equipment in the revision."""
    data = {name: getattr(tables, name, []) for name in
            ("nodes", "pipes", "nozzles", "fittings", "equipment", "pumps", "valves", "meta", "unresolved")}
    return hashlib.sha256(json.dumps(data, ensure_ascii=False, sort_keys=True,
                                    allow_nan=False).encode()).hexdigest()


def _number(value, title, *, minimum=None):
    try:
        value = float(value)
    except (ValueError, TypeError):
        raise ValueError(f"{title}: 숫자가 필요합니다.") from None
    if not math.isfinite(value) or (minimum is not None and value < minimum):
        raise ValueError(f"{title}: 값의 범위를 확인하세요.")
    return value


def default_source(tables: Any) -> str:
    """A single pump outlet, otherwise one input node; never guess among sources."""
    pumps = getattr(tables, "pumps", []) or []
    if len(pumps) == 1:
        return str(pumps[0]["out"])
    sources = [str(n["label"]) for n in tables.nodes if n.get("io_node") == "Input"]
    return sources[0] if len(sources) == 1 and not pumps else ""


def prepare(tables: Any, options: dict, *, library: dict | None = None) -> tuple[Network, Settings, list[str]]:
    """Use authoritative SLF bores/K values and counted fitting losses once.

Unknown pipe roles use the stricter branch limit. Unknown loss elements block
calculation unless a user explicitly provides that pipe's total equivalent
length; such pipes are locked because their size-dependent loss is unknown.
"""
    if library is None:
        from routes.module_f.network_edit import catalog
        library = catalog()
    drawing_first = any((p.get('bore_provenance') or {}).get('policy') == 'drawing_first_v1' for p in tables.pipes)
    scope = options.get('diameter_scope', 'unspecified' if drawing_first else 'enlarge')
    if scope not in ('unspecified','all','enlarge'):
        raise ValueError('역산 관경 범위를 확인하세요.')
    review = fitting_review(tables)
    if review.count:
        spots = ' · '.join(f"노드 {r.get('node_label', '?')} / 배관 {r.get('pipe_label', r.get('pipe', '?'))}"
                           for r in review.pending[:8])
        if review.unlocated_count:
            spots += f" · 위치 정보 없는 미확정 {review.unlocated_count}건(표 재확정 필요)"
        raise ValueError(f"부속 판정 미확정 {review.count}건: {spots} — 해당 노드·배관에 부속을 지정한 뒤 결합망을 다시 불러오세요. 누락된 손실을 0으로 계산하지 않습니다.")
    settings = Settings(**{key: options[key] for key in Settings.__dataclass_fields__ if key in options})
    settings.validate()
    max_dn = _number(options.get("max_dn", 200), "최대 호칭경", minimum=1)
    min_flow = _number(options.get("min_flow_lpm", 80), "헤드 최소 유량", minimum=.001)
    min_bar = _number(options.get("min_pressure_bar", 1), "헤드 최소 압력", minimum=.001)
    max_bar = _number(options.get("max_head_pressure_bar", 12), "헤드 최대 압력", minimum=.001)
    if getattr(tables, "valves", None):
        raise ValueError("독립 제어밸브 링크(PRV 등)는 아직 지원하지 않습니다. 임의의 일반 배관으로 바꾸지 않습니다.")
    pumps = getattr(tables, "pumps", []) or []
    if len(pumps) > 1:
        raise ValueError("다중 펌프 링크는 아직 지원하지 않습니다. 단일 토출 경계의 검토망이 필요합니다.")
    source = str(options.get("source") or default_source(tables))
    if not source:
        raise ValueError("공급 경계 절점을 지정하세요.")
    if pumps and source != str(pumps[0]["out"]):
        raise ValueError("펌프가 있는 망은 해당 펌프의 토출 절점을 공급 경계로 지정하세요.")
    inputs = {str(n["label"]) for n in tables.nodes if n.get("io_node") == "Input"}
    ignored = {str(pumps[0]["in"])} if pumps else set()
    if inputs - ignored - {source}:
        raise ValueError("여러 공급 경계를 단일 압력으로 대신할 수 없습니다.")
    warnings = ["정상상태 물 계산 · Hazen–Williams · 20°C. 법규 적합/최종 설계 확정이 아닙니다.",
                "현재 선택된 작동 헤드 시나리오만 검토합니다. 다른 작동 구역의 최불리성은 별도 검토하세요.",
                ("미지정 관경만 역산하며 도면·사용자 관경은 잠급니다." if scope=='unspecified' else
                 "전체 관경을 후보 규격부터 다시 검토합니다. 전역 최적해를 보장하지 않습니다." if scope=='all' else
                 "관경은 기존값 이상으로만 확대합니다. 최소 비용·전역 최적해를 보장하지 않습니다.")]
    for row in review.resolved:
        warnings.append(f"노드 {row['node_label']} / 배관 {row.get('pipe_label', row.get('pipe'))}: "
                        f"사용자 지정 부속 {', '.join(row['fitting_ids'])}으로 미확정 판정 해소 · 손실 반영")
    if settings.discharge_velocity_head_m == 0:
        warnings.append("펌프 토출 속도수두를 0m로 입력한 추정 양정입니다. 실제 토출 유속 v²/(2g)를 반영하고 제조사 곡선·효율·NPSH·정지압을 별도 확인하세요.")
    if pumps:
        warnings.append("펌프 곡선은 사용하지 않습니다. 펌프 토출의 지정압력/필요압력 경계로 검토하며, 흡입측 전수두는 직접 입력값입니다.")
    refs = {str(p[k]) for p in tables.pipes for k in ("in", "out")}
    if ignored & refs:
        raise ValueError("펌프 흡입측 배관이 포함되어 있습니다. 이번 계산은 토출 이후 망만 지원합니다. 흡입망 손실을 누락한 채 계산하지 않습니다.")
    nodes = tuple(Node(str(n["label"]), _number(n.get("elevation"), f"절점 {n.get('label')} 표고"))
                  for n in tables.nodes if str(n["label"]) not in ignored)
    pipe_ids = {str(p["label"]) for p in tables.pipes}
    edits = options.get("pipes") or {}
    if not isinstance(edits, dict) or set(edits) - pipe_ids:
        raise ValueError("개별 배관 설정에 현재 망에 없는 라벨이 있습니다.")
    fitting_by_pipe = {label: [] for label in pipe_ids}
    equipment_by_pipe = {label: [] for label in pipe_ids}
    for rows, dest in ((tables.fittings, fitting_by_pipe), (tables.equipment, equipment_by_pipe)):
        for row in rows:
            pid = str(row.get("pipe"))
            if pid not in dest:
                raise ValueError(f"손실 요소의 배관 {pid}가 없습니다.")
            dest[pid].append(row)
    # The merge builder uses the native SDF valve names, whereas the plan
    # editor exposes library IDs. These mean the same library items (see
    # kfp_sdf_converter's native fitting-name mapping); do not discard losses.
    aliases = {"gate":"VALVE_GATE", "check":"VALVE_SWING_CHECK",
               "butterfly":"VALVE_BUTTERFLY", **library["fitting_aliases"]}
    fit_defs = {r["id"]: r for r in library["fittings"]}
    pipes = []
    unknown_roles = 0
    for row in tables.pipes:
        label = str(row["label"])
        edit = edits.get(label) or {}
        if not isinstance(edit,dict):
            raise ValueError(f"배관 {label}: 개별 설정 형식 오류")
        if str(row.get("status", "Normal")).lower() not in {"normal", "open", "1"}:
            raise ValueError(f"배관 {label}: 닫힘/제어 상태는 이번 역산에서 지원하지 않습니다.")
        unknown = drawing_first and (row.get('bore_provenance') or {}).get('policy')=='drawing_first_v1' and not row.get('dia')
        dn = 0 if unknown else _number(row.get("dia"), f"배관 {label} 호칭경", minimum=.001)
        schedule = str(row.get("type"))
        catalog = sorted([p for p in library["pipes"] if p["type"] == schedule], key=lambda p:p["dia"])
        current = next((p for p in catalog if abs(p["dia"]-dn)<.001), None)
        if current is None and not unknown:
            raise ValueError(f"배관 {label}: {schedule} {dn:g}A의 SLF 실제 내경이 없습니다.")
        if not unknown and row.get("inner_mm") is not None and abs(float(row["inner_mm"])-current["inner_mm"]) > .01:
            raise ValueError(f"배관 {label}: 지정 내경과 SLF 내경이 다릅니다. 규격을 먼저 확인하세요.")
        role = str(edit.get("role") or row.get("pipe_role") or "unknown")
        unknown_roles += role == "unknown"
        override = edit.get("equivalent_m")
        if override is not None:
            override = _number(override, f"배관 {label} 전체 등가길이", minimum=0)
            warnings.append(f"{label}: 사용자 지정 전체 등가길이 {override:g}m, 관경 잠금 적용")
        locked = bool(edit.get("locked", False)) or override is not None or (scope=='unspecified' and not unknown)
        if unknown and locked:
            raise ValueError(f'배관 {label}: 미지정 관경을 잠글 수 없습니다. 관경별 손실 자료와 잠금 설정을 확인하세요.')
        eq_issues = []

        def equivalent(candidate_dn):
            if override is not None:
                return override
            total = 0.
            fits = fitting_by_pipe[label]
            # Pipe eq_len is the cached fitting sum, NOT another independent loss.
            if not fits and float(row.get("eq_len") or 0) > 0:
                raise ValueError("부속 상세 없이 합계만 있습니다. 전체 등가길이를 확인·입력하세요.")
            for f in fits:
                kind = str(f.get("type"))
                if kind in {"tee-run", "cross-run"}:
                    continue
                ident = aliases.get(kind, kind)
                value = (fit_defs.get(ident, {}).get("lengths") or {}).get(str(int(candidate_dn)))
                if value is None:
                    raise ValueError(f"{kind} {candidate_dn:g}A 등가길이 없음")
                total += _number(value, f"{kind} 등가길이", minimum=0)*_number(f.get("count",1), "부속 개수",minimum=0)
            for e in equipment_by_pipe[label]:
                if e.get('editor_library'):
                    ident = str(e['editor_library'])
                    values = fit_defs.get(ident, {}).get('lengths') or {}
                    value = values.get(str(int(candidate_dn)))
                    if value is None:
                        raise ValueError(f"{ident} {candidate_dn:g}A 등가길이 없음")
                    count = _number(e.get('count', 1), '부속 개수', minimum=1)
                    loss = _number(value, f'{ident} 등가길이', minimum=0) * count
                    if candidate_dn == dn and abs(_number(e.get('eq_len'), '저장 등가길이', minimum=0)-loss) > .001:
                        raise ValueError(f"{ident}: 저장된 손실과 라이브러리 값이 다릅니다. 부속 또는 전체 등가길이를 확인하세요.")
                    total += loss
                    continue
                value = _number(e.get("eq_len"), f"기기 {e.get('desc')} 등가길이", minimum=0)
                if value == 0 and not e.get("eq_len_src"):
                    raise ValueError(f"기기 {e.get('desc')}의 0m 손실이 미확정입니다. 전체 등가길이를 확인·입력하세요.")
                total += value
            return total

        candidates = []
        for spec in catalog:
            d = spec["dia"]
            if (locked and d!=dn) or (not locked and d>max_dn and not (scope=='enlarge' and d==dn)) or (scope=='enlarge' and d<dn):
                continue
            try:
                eq = equivalent(d)
            except ValueError as exc:
                if d == dn:
                    raise ValueError(f"배관 {label}: {exc}") from exc
                eq_issues.append(f"{d:g}A"); continue
            candidates.append(Size(d, spec["inner_mm"], eq))
        if eq_issues:
            warnings.append(f"{label}: 손실 자료 없는 관경 후보 제외 ({', '.join(eq_issues)})")
        if any(f.get("loss_status") or f.get("flow_direction") == "solver_reference" for f in fitting_by_pipe[label]):
            warnings.append(f"{label}: 루프 분류티 손실은 현재 귀속 배관의 잠정 등가길이 모델입니다. 실제 분류/합류 계수 재검토 필요")
        if any(not e.get('editor_library') for e in equipment_by_pipe[label]) and not locked:
            # Existing equipment eq lengths generally do not carry a diameter
            # response. Keep that host fixed rather than reuse a wrong-sized AV.
            if unknown:
                raise ValueError(f'배관 {label}: 미지정 배관의 기기 손실에 관경별 자료가 없습니다. 관경을 먼저 정의하세요.')
            locked = True
            candidates = [size for size in candidates if size.nominal_mm==dn]
            warnings.append(f"{label}: 기기 손실의 관경별 자료가 없어 관경 잠금 적용")
        pipes.append(Pipe(label,str(row["in"]),str(row["out"]),
            _number(row.get("length"), f"배관 {label} 길이",minimum=.000001),
            _number(row.get("c"), f"배관 {label} C",minimum=.001),tuple(candidates),locked=locked,role=role))
    if unknown_roles:
        warnings.append(f"역할 미지정 {unknown_roles}개 배관에는 보수적으로 가지배관 유속 제한을 적용합니다.")
    active = options.get("active_nozzles")
    all_heads = {str(h["label"]):h for h in tables.nozzles}
    if len(all_heads) != len(tables.nozzles):
        raise ValueError("노즐 라벨 중복")
    if active is None:
        active = [key for key,h in all_heads.items() if str(h.get("status", "1")) == "1"]
    if not isinstance(active,list) or not active or set(map(str,active)) - all_heads.keys():
        raise ValueError("현재 망의 작동 노즐 라벨을 하나 이상 지정하세요.")
    nozzle_lib = {h["id"]:h for h in library["nozzles"]}
    nozzles = []
    for label in dict.fromkeys(map(str,active)):
        h = all_heads[label]
        spec = nozzle_lib.get(h.get("lib"))
        if spec is None:
            raise ValueError(f"헤드 {label}: SLF 노즐 K 값이 없습니다.")
        if h.get('definition_policy')=='drawing_first_v1':
            from src.pipenet_converter.sdf.nozzle_overrides import corrected_spec
            import math
            spec=corrected_spec(h,dict(k_si=spec['k_factor_si']/60000/math.sqrt(100000),
                min_p_pa=spec['min_bar']*100000,max_p_pa=spec['max_bar']*100000))
        nozzles.append(Nozzle(label, str(h["in"]),spec["k_factor_si"],min_flow,
            max(min_bar,spec["min_bar"]),min(max_bar,spec["max_bar"]) if spec["max_bar"]>0 else max_bar))
    network = Network(nodes,tuple(pipes),tuple(nozzles),source)
    network.validate()
    return network, settings, warnings


def input_summary(tables: Any) -> dict:
    """Small editor data, not the large CAD session."""
    review = fitting_review(tables)
    return dict(source=default_source(tables), fingerprint=fingerprint(tables),
        fitting_review=dict(pending_count=review.count,pending=list(review.pending),
                            resolved=list(review.resolved),unlocated_count=review.unlocated_count),
        pipes=[dict(label=str(p['label']),a=str(p['in']),b=str(p['out']),dia=p.get('dia'),
                    type=p.get('type'),role=p.get('pipe_role') or 'unknown') for p in tables.pipes],
        nozzles=[dict(label=str(h['label']),node=str(h['in']),lib=h.get('lib'),
                      active=str(h.get('status','1'))=='1') for h in tables.nozzles])

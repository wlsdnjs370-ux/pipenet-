"""Read-only fitting annotations for plan and merged calculation networks.

Types come from the existing fitting table, never from the pruned display degree.
Display arms/positions are annotations only: no pipe, length or loss is written.
Library lookup is a reference alongside the actually stored pipe loss; this module
excludes straight-through tees and never silently resolves missing values.
"""
from __future__ import annotations

from collections import defaultdict
import math
from typing import Any

from src.pipenet_converter.graph.fitting_policy import without_straight_tees
from src.pipenet_converter.graph.fitting_review import fitting_review


# 겹친 노드 판정 — convert.main_walk.JOINT_M(0.01 m)과 같은 값을 board mm 로.
#   화면 표시용이라 엔진을 import 하지 않고 수만 맞춘다.
JOINT_MM = 10.0

NAMES = {"tee": "분류티", "elbow": "90° 엘보",
         "elbow-45": "45° 엘보", "alarm_valve": "알람밸브"}


def _unit(dx: float, dy: float) -> list[float] | None:
    length = math.hypot(dx, dy)
    return [dx / length, dy / length] if length > 1e-9 else None


def _shape(kind: str) -> str:
    if "cross" in kind.lower():
        return "cross"
    if "tee" in kind.lower():
        return "tee"
    if "elbow" in kind.lower():
        return "elbow"
    if "valve" in kind.lower():
        return "valve"
    return "other"


def build_inspection(tables: Any, got: dict, keys: dict, nodes: list[dict], *,
                     board: Any = None, transform: dict | None = None,
                     edited: bool = False, base_got: dict | None = None,
                     changed_nodes: set[str] | None = None) -> dict:
    """Return per-fitting cards and projected glyphs using existing identities.

    ``nodes`` are display coordinates; board X/Y are mm. Only unit display
    directions cross this boundary. ``eq_m`` is m per fitting, while pipe totals
    are the stored calculation values. Glyph radii use display units, not pixels.
    Unknown positions stay unplaced; omitted branches are never invented.
    """
    from services.cad_import.design.fitting import (
        FITTING_LIB_ID, load_equivalent_lengths, resolve_eq_len)
    from services.cad_import.design.restrict import board_port_neighbors

    lib = load_equivalent_lengths()
    at = {str(n["label"]): n for n in nodes}
    pipes = {str(p["label"]): p for p in tables.pipes}
    incident: dict[str, list[str]] = defaultdict(list)
    upstream: dict[str, list[str]] = defaultdict(list)
    for lab, p in pipes.items():
        upstream[str(p["out"])].append(str(p["in"]))
        for end in (str(p["in"]), str(p["out"])):
            incident[end].append(lab)
    pts = list(getattr(board, "pts", None) or [])
    adj: dict[int, set[int]] = defaultdict(set)
    for a, b in (getattr(board, "edges", None) or []):
        a, b = int(a), int(b)
        if 0 <= a < len(pts) and 0 <= b < len(pts):
            adj[a].add(b)
            adj[b].add(a)
    nref = {str(k): int(v) for k, v in (got.get("node_ref") or {}).items()}
    # Separate from editor/nozzle identity: the head and its new takeoff share
    # an original plan point but are different calculation nodes.
    fref = {**nref, **{str(k): int(v) for k, v in
                     (got.get("fitting_node_ref") or {}).items()}}
    phys = {str(k): int(v) for k, v in (got.get("phys") or {}).items()}
    # 계산 표의 평면 좌표·표고 — 세로관(같은 평면 자리, 다른 표고)을 알아보는 데 쓴다.
    plan_xyz = {str(n["label"]): (float(n.get("x") or 0.0), float(n.get("y") or 0.0),
                                  float(n.get("elevation") or 0.0))
                for n in (getattr(tables, "nodes", None) or [])}
    # 호 갈래(`arc_junctions`)의 주배관 접속점·꼭대기 — 그 자리 티의 옆구멍은 세로관이다.
    lab_of_nid = {str(v): str(k) for k, v in (keys.get("nid") or {}).items()}
    arc_tee: set[str] = set()
    for m_nid, rec in (((got.get("kfp") or {}).get("arc_junctions")) or {}).items():
        top_nid = rec.get("top") if isinstance(rec, dict) else None
        for nid in (m_nid, top_nid):
            if nid is not None and str(nid) in lab_of_nid:
                arc_tee.add(lab_of_nid[str(nid)])

    def unchanged(label: str) -> bool:
        if label in (changed_nodes or ()):
            return False
        if not edited:
            return True
        nid = str(keys.get("nid", {}).get(label, ""))
        current = (got.get("kfp") or {}).get("nodes_meta_runtime") or {}
        before = ((base_got or {}).get("kfp") or {}).get("nodes_meta_runtime") or {}
        xyz = (current.get(nid) or {}).get("coords")
        return bool(xyz and xyz == (before.get(nid) or {}).get("coords"))

    def radius(label: str | None, pl: str) -> float:
        lengths = []
        for name in incident.get(label, ()) if label else [pl]:
            p = pipes[name]
            a, b = at.get(str(p["in"])), at.get(str(p["out"]))
            if a and b:
                length = math.hypot(b["x"] - a["x"], b["y"] - a["y"])
                if length > 1e-9:
                    lengths.append(length)
        cap = 200.0 * abs(float((transform or {}).get("k", 1.0)))
        return min(cap, min(lengths) * .22) if lengths else cap

    def original_arms(vid: int | None, skip=frozenset()) -> list[list[float]]:
        if vid is None or not transform or not (0 <= vid < len(pts)):
            return []
        result = []
        for other in board_port_neighbors(pts, {vid: adj.get(vid, set()) - set(skip)}, vid):
            dx, dy = pts[other][0] - pts[vid][0], pts[other][1] - pts[vid][1]
            if transform.get("iso"):
                dx, dy = ((dx - dy) * transform["cos30"],
                          (dx + dy) * transform["sin30"])
            u = _unit(dx, dy)
            if u and not any(sum(a*b for a, b in zip(u, v)) > .98 for v in result):
                result.append(u)
        return result

    def arc_run_arm(label: str) -> tuple[list[float] | None, str | None]:
        """호 갈래 티의 직선(가로) 팔 방향과 세로관 반대 끝 — 세로관 하나 + 가로 하나일 때만."""
        n, here3 = at.get(label), plan_xyz.get(label)
        if not n or not here3:
            return None, None
        riser_end, run = None, []
        for pl in incident[label]:
            p = pipes[pl]
            end = str(p["out"]) if str(p["in"]) == label else str(p["in"])
            there3, o = plan_xyz.get(end), at.get(end)
            if not there3 or not o:
                return None, None
            if (math.hypot(here3[0] - there3[0], here3[1] - there3[1]) <= 1e-6
                    and abs(here3[2] - there3[2]) > 1e-9):
                if riser_end is not None:
                    return None, None
                riser_end = end
            else:
                run.append(_unit(o["x"] - n["x"], o["y"] - n["y"]))
        if riser_end is None or len(run) != 1 or not run[0]:
            return None, None
        return run[0], riser_end

    def plan_dirs(vid: int | None) -> list[list[float]]:
        """원본 도면에서 이 자리가 실제로 뻗는 방향(표시 단위벡터). 겹친 노드(≤ JOINT_MM)는
        건너가 그 너머 배관의 방향으로 잰다 — 0.04 mm 연결관의 방향은 잡음이다."""
        if vid is None or not transform or not (0 <= vid < len(pts)):
            return []
        near = {o for o in adj.get(vid, ()) if 0 <= o < len(pts)
                and math.dist(pts[vid][:2], pts[o][:2]) <= JOINT_MM}
        far = {o for o in adj.get(vid, ()) if o not in near and 0 <= o < len(pts)}
        for q in near:
            far |= {o for o in adj.get(q, ()) if o != vid and o not in near and 0 <= o < len(pts)
                    and math.dist(pts[vid][:2], pts[o][:2]) > JOINT_MM}
        result = []
        for other in sorted(far):
            dx, dy = pts[other][0] - pts[vid][0], pts[other][1] - pts[vid][1]
            if transform.get("iso"):
                dx, dy = ((dx - dy) * transform["cos30"], (dx + dy) * transform["sin30"])
            u = _unit(dx, dy)
            if u:
                result.append(u)
        return result

    def node_geometry(label: str) -> tuple[list[list[float]], int | None, str]:
        n = at.get(label)
        arms = []
        if n:
            for pl in incident[label]:
                p = pipes[pl]
                end = str(p["out"]) if str(p["in"]) == label else str(p["in"])
                o = at.get(end)
                u = _unit(o["x"] - n["x"], o["y"] - n["y"]) if o else None
                if u:
                    arms.append(u)
        nid = str(keys.get("nid", {}).get(label, ""))
        vid = fref.get(nid)
        degree = phys.get(nid)
        # An edited graph may have moved or split this source node. Retain the
        # recorded fitting type, but never pretend its old directions are current.
        stable = unchanged(label) and all(
            unchanged(str(pipes[pl]["out"] if str(pipes[pl]["in"]) == label
                          else pipes[pl]["in"])) for pl in incident[label])
        confirmed_ports = (got.get('kfp') or {}).get('arc_ports',{}).get(nid)
        if stable and confirmed_ports and transform:
            projected=[]
            for vx,vy,vz in confirmed_ports:
                dx,dy=vx*1000*float(transform.get('k',1)),vy*1000*float(transform.get('k',1))
                if transform.get('iso'):
                    dx,dy=(dx-dy)*transform['cos30'],(dx+dy)*transform['sin30']+vz*float(transform.get('lift',1000))
                u=_unit(dx,dy)
                if u: projected.append(u)
            return projected,len(confirmed_ports),'원본 호의 포트 분리 + 검증된 입체 접속'
        # ★[괄호 교차 꼭대기 · 오너 2026-09-22] 겹친 노드(≤ JOINT_M) 사이의 평면
        #   연결을 전개가 이 노드의 세로관으로 세웠으면, 그 팔은 **세로 팔로 이미
        #   그려져 있다** — «빠진 방향» 후보가 아니다. 이것을 후보에 두면 빠진 팔
        #   1개(가지치기로 잘린 가지 반쪽)에 후보가 2개가 되어 아래 «모호하면 안
        #   그린다» 에 걸리고, 분류티가 ㄱ자(팔 2)로 그려져 엘보처럼 보였다.
        #   부속 종류·등가길이는 부속표 그대로다 — 기호의 팔만 바뀐다.
        stood_up = set()
        here3 = plan_xyz.get(label)
        for pl in incident[label]:
            p = pipes[pl]
            end = str(p["out"]) if str(p["in"]) == label else str(p["in"])
            there3 = plan_xyz.get(end)
            if not here3 or not there3 or vid is None:
                continue
            if (math.hypot(here3[0] - there3[0], here3[1] - there3[1]) > 1e-6
                    or abs(here3[2] - there3[2]) <= 1e-9):
                continue                                  # 세로관이 아니다
            v_end = fref.get(str(keys.get("nid", {}).get(end, "")))
            if (v_end is not None and 0 <= int(v_end) < len(pts) and 0 <= vid < len(pts)
                    and int(v_end) in adj.get(vid, ())
                    and math.dist(pts[vid][:2], pts[int(v_end)][:2]) <= JOINT_MM):
                stood_up.add(int(v_end))
        originals = original_arms(vid, stood_up) if stable else []
        omitted = [u for u in originals
                   if not any(sum(a*b for a, b in zip(u, v)) > .94 for v in arms)]
        missing = max(0, (degree if degree is not None else len(originals)) - len(arms))
        # Elevation expansion can replace an existing plan arm with a vertical
        # arm. It must not be drawn twice. Add omitted directions only when the
        # source correspondence is unambiguous, never truncate arbitrary arms.
        if missing and len(omitted) == missing:
            arms.extend(omitted)
        basis = ("이 위치의 방향 변경 · 현재 연결 방향만 표시, 생략 방향 재확인" if not stable else
                 "원본 연결 방향 + 현재 계산망" if originals else
                 "현재 계산망 · 생략 방향 미확인")
        return arms, degree, basis

    fits = without_straight_tees(tables.fittings)

    records: list[dict] = []
    unresolved = getattr(tables, "unresolved", None) or {}
    for i, row in enumerate(fits):
        pl, kind = str(row.get("pipe")), str(row.get("type") or "unknown")
        p = pipes.get(pl)
        if not p:
            continue
        lab = str(row.get("node") or p["in"])
        anchor = at.get(lab)
        arms, degree, basis = node_geometry(lab)
        count = int(row.get("count") or 1)
        source_node = fref.get(str(keys.get("nid", {}).get(lab, "")))
        eq, why = resolve_eq_len(kind, p.get("dia"), lib=lib)
        for ov in unresolved.get("applied", []):
            if (ov.get("what") == "eq_len" and str(ov.get("pipe_label", ov.get("pipe"))) == pl
                    and ov.get("kind") == kind and ov.get("dia") == p.get("dia")):
                eq, why = ov.get("m"), "직접 입력: " + str(ov.get("note") or "사유 미기록")
        if why == "라이브러리":
            why = "fittings_library_v3.json / " + str(FITTING_LIB_ID.get(kind, kind))
        shape = _shape(kind)
        # ★[호 갈래 티 · 오너 2026-09-22] 호 갈래의 주배관 접속점·꼭대기 티는 세로관이
        #   옆구멍이고 나머지 두 팔이 한 직선이다(주배관 양쪽 / 가지관 양쪽). 위의 원본
        #   연결 되찾기가 모호하거나(겹친 노드·짧은 토막) 꼭대기를 엔진이 새로 만들어
        #   원본 짝이 없으면 팔이 둘뿐이라 ㄱ자(엘보처럼)로 그려졌다. 직선 팔의 맞은편에
        #   **원본 도면 배관이 실제로 있을 때만** 셋째 팔로 그린다 — 꼭대기에 원본 짝이
        #   없으면 세로관 아래 주배관 접속점의 원본 연결로 확인한다. 추측으로 긋지 않는다.
        #   부속 종류·등가길이는 부속표 그대로다.
        if shape == "tee" and len(arms) == 2 and lab in arc_tee:
            run_u, riser_end = arc_run_arm(lab)
            if run_u is not None and unchanged(lab) and unchanged(riser_end):
                opp = [-run_u[0], -run_u[1]]
                v_here = fref.get(str(keys.get("nid", {}).get(lab, "")))
                v_base = (v_here if v_here is not None else
                          fref.get(str(keys.get("nid", {}).get(riser_end, ""))))
                if (not any(sum(a*b for a, b in zip(opp, v)) > .94 for v in arms)
                        and any(sum(a*b for a, b in zip(opp, v)) > .94 for v in plan_dirs(v_base))):
                    arms = arms + [opp]
                    basis = "현재 계산망 + 호 갈래 티의 직선 맞은편 (원본 도면 배관으로 확인)"
        need = {"tee": 3, "cross": 4, "elbow": 2}.get(shape, 0)
        symbolic = len(arms) != need if need else True
        review = row.get('flow_direction') == 'solver_reference'
        flow = ([str(p['in']), str(p['out'])] if review else
                [upstream[lab][0] if len(upstream[lab]) == 1 else "상류 미확정",
                 lab, str(p["out"])] if lab else [str(p["in"]), "중간 부속", str(p["out"])])
        records.append(dict(id=f"f{i}", pipe=pl, node=lab, kind=kind,
            name=NAMES.get(kind, kind), shape=shape, count=count,
            x=anchor["x"] if anchor else None, y=anchor["y"] if anchor else None,
            arms=arms, glyph_radius=radius(lab, pl), symbolic=symbolic, geometry_source=basis,
            original_degree=degree, source_node=source_node,
            current_degree=len(incident[lab]) if lab else 2,
            material=p.get("type"), dia=p.get("dia"), eq_m=eq,
            eq_total_m=eq*count if eq is not None else None,
            eq_source=why or "해당 종류·관경의 등가길이 미확정",
            pipe_eq_m=p.get("eq_len"), origin="부속 입력표", flow_path=flow,
            flow_label="손실 귀속 배관 · 유향 미확정" if review else "계산 경로",
            loss_status=row.get('loss_status'),
            note=("실제 접속점에 표시 · 티 손실은 가지 포트의 검토용 값이며 실제 유향·손실은 수리계산 확인 필요"
                  if review else "종류는 원본 정보를 반영한 부속표 기준 · 축약된 선 모양으로 재판정하지 않음")))

    # Explicit point fittings/valves added with the direct editor live in the
    # equipment table, not the native fitting table. They must be visible too.
    for i, row in enumerate(tables.equipment):
        pl = str(row.get("pipe"))
        p = pipes.get(pl)
        if not p:
            continue
        a, b = at.get(str(p["in"])), at.get(str(p["out"]))
        if not a or not b:
            continue
        t = max(0.0, min(1.0, float(row.get("rel_pos", .5))))
        kind = str(row.get("editor_library") or row.get("lib") or
                   ("alarm_valve" if row.get("desc") == "A/V" else "equipment"))
        count = int(row.get("count") or 1)
        value = row.get("eq_len")
        lab = row.get("editor_node") or (str(p["in"]) if t == 0 else str(p["out"]) if t == 1 else None)
        records.append(dict(id=f"e{i}", pipe=pl, node=lab, kind=kind,
            name=row.get("desc") or NAMES.get(kind, kind), shape=_shape(kind),
            count=count, x=a["x"]+(b["x"]-a["x"])*t, y=a["y"]+(b["y"]-a["y"])*t,
            arms=[], glyph_radius=radius(lab, pl), symbolic=True,
            geometry_source="기기표 설치 위치 · 연결 포트 형상 미확인",
            original_degree=None, current_degree=len(incident[str(lab)]) if lab else None,
            material=p.get("type"), dia=p.get("dia"),
            eq_m=float(value)/count if value is not None else None, eq_total_m=value,
            eq_source="기기표 저장값 / " + str(row.get("editor_library") or row.get("lib") or "직접 입력·기존 값"),
            pipe_eq_m=p.get("eq_len"), origin="기기 입력표", note="배관 실제 길이와 별도로 적용"))

    # Unclassified locations are visible, but never drawn as a guessed elbow/tee.
    for i, row in enumerate(fitting_review(tables).pending):
        pl = str(row.get("pipe_label", row.get("pipe")))
        p = pipes.get(pl)
        if not p:
            continue
        lab = (None if row.get("where") == "구간 내부 다중 접속"
               else str(row.get("node_label") or p["in"]))
        n = at.get(lab)
        if n or lab is None:
            records.append(dict(id=f"u{i}", pipe=pl, node=lab, kind="unresolved",
                name="부속 판정 미확정", shape="other", count=row.get("n", 1),
                x=n["x"] if n else None, y=n["y"] if n else None,
                arms=[], glyph_radius=radius(lab, pl), symbolic=True,
                geometry_source="미확정 부속 위치", original_degree=None,
                current_degree=len(incident[lab]), material=p.get("type"), dia=p.get("dia"),
                eq_m=None, eq_total_m=None, eq_source="종류 확인 필요", origin="미해결 목록",
                note=row.get("reason") or "원본 도면과 부속 판정 근거를 확인하세요."))
    return {"fittings": records, "edited": edited,
            "note": "표시 전용 · 배관 실제 길이/연결/계산 손실은 변경하지 않음"}


def build_merged_inspection(tables: Any, plan: dict, nodes: list[dict], *,
                            offset: int, board: Any = None,
                            transform: dict | None = None,
                            plan_editor: dict | None = None,
                            merge_editor: dict | None = None) -> dict:
    """Project authoritative combined fittings, preserving original plan ports.

    Node IDs shift at merge; pipe IDs may be renamed on collision. Neither
    projected degree nor the shape of a schematic riser reclassifies a fitting.
    Changed merged nodes use only their confirmed current connection directions.
    """
    from copy import copy, deepcopy
    from routes.module_f.merge import _shift

    keys = plan.get("keys") or {}
    mapped = {"nid": {_shift(lab, offset): nid
                      for lab, nid in (keys.get("nid") or {}).items()}}
    got = plan.get("got") or {}
    view_tables = copy(tables)
    source = plan.get("tables")
    # Match provenance by endpoints, not a reused pipe name (r1/P1 collisions).
    pipes = {(str(p["in"]), str(p["out"])): str(p["label"])
             for p in tables.pipes}
    rename = {str(p["label"]): pipes.get((_shift(p["in"], offset),
                                          _shift(p["out"], offset)))
              for p in (getattr(source, "pipes", None) or [])}
    unresolved = deepcopy(getattr(source, "unresolved", None) or {})
    for group in ("kind_items", "applied"):
        kept = []
        for row in unresolved.get(group, []):
            name = rename.get(str(row.get("pipe_label", row.get("pipe"))))
            if name is None:
                continue
            row["pipe_label"] = name
            if row.get("node_label") is not None:
                row["node_label"] = _shift(row["node_label"], offset)
            kept.append(row)
        unresolved[group] = kept
    # New combined tables carry the exact source issues through renaming. Keep
    # the mapped fallback for legacy tables that predate this metadata.
    view_tables.unresolved = deepcopy(getattr(tables, 'unresolved', None) or unresolved)
    base = ((merge_editor or {}).get("base_object") or {}).get("combined")
    before = {str(n["label"]): n for n in (getattr(base, "nodes", None) or [])}
    changed = {str(n["label"]) for n in tables.nodes
               if base is not None and any(n.get(k) != before.get(str(n["label"]), {}).get(k)
                                           for k in ("x", "y", "elevation"))}
    return build_inspection(view_tables, got, mapped, nodes, board=board,
                            transform=transform, changed_nodes=changed,
                            edited=bool(plan.get("editor_modified")),
                            base_got=((plan_editor or {}).get("base_object") or {}).get("got"))

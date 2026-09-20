"""KFP edit compatibility without silently changing calculation geometry.

Coordinates and lengths are metres. The default export only sets the editor
grid. A separately requested legacy editing copy uses the Solver's 1 cm grid
and reports every changed pipe length. No Flask or editor runtime is needed.
"""
from __future__ import annotations

import copy
import math
from dataclasses import asdict, dataclass, field
from typing import Any

TOL_M = 1e-6
SOLVER_GRID_M = 0.01
PRECISE_GRID_M = 1e-6
MAX_CONTACT_CHECKS = 250_000
Coord = tuple[float, float, float]


@dataclass
class EditReport:
    """Inspection result and explicit differences from the calculation copy."""

    mode: str
    nodes: int = 0
    pipes: int = 0
    off_grid_nodes: list[str] = field(default_factory=list)
    non_axis_pipes: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    changed_nodes: int = 0
    length_changes: list[dict[str, Any]] = field(default_factory=list)
    max_coordinate_change_mm: float = 0.0
    max_length_change_mm: float = 0.0
    total_length_change_mm: float = 0.0

    def to_dict(self) -> dict:
        """Return a JSON-ready report with a clear acceptance flag."""
        return {"ok": not self.errors, **asdict(self)}


def _positions(data: dict, report: EditReport) -> dict[str, Coord]:
    meta = data.get("nodes_meta_runtime")
    legacy = data.get("nodes")
    raw = meta if isinstance(meta, dict) else legacy
    if not isinstance(raw, dict) or not raw:
        report.errors.append("KFP에 노드 좌표가 없습니다.")
        return {}
    coords = {}
    for name, item in raw.items():
        try:
            values = item.get("coords") if isinstance(meta, dict) else item
            if len(values) != 3:
                raise ValueError("coordinate dimension")
            xyz = tuple(float(v) for v in values)
            if not all(math.isfinite(v) for v in xyz):
                raise ValueError("nonfinite coordinate")
            coords[str(name)] = xyz
        except (TypeError, ValueError, AttributeError):
            report.errors.append(f"노드 {name}: 유효한 XYZ 좌표가 아닙니다.")
    return coords


def _axis(a: Coord, b: Coord) -> int | None:
    axes = [i for i in range(3) if abs(a[i] - b[i]) > TOL_M]
    return axes[0] if len(axes) == 1 else None


def _box(a: Coord, b: Coord) -> tuple[Coord, Coord]:
    return tuple(min(x, y) for x, y in zip(a, b)), tuple(max(x, y) for x, y in zip(a, b))


def _touch(a: tuple[Coord, Coord], b: tuple[Coord, Coord]) -> bool:
    return all(a[0][i] <= b[1][i] + TOL_M and b[0][i] <= a[1][i] + TOL_M
               for i in range(3))


def _new_contacts(old: dict, new: dict, edges: dict, report: EditReport) -> None:
    """Reject new contacts between axis-aligned pipes, with bounded work."""
    old_boxes = {pid: _box(old[a], old[b]) for pid, (a, b) in edges.items()}
    boxes = {pid: _box(new[a], new[b]) for pid, (a, b) in edges.items()}
    # Sweep along the widest coordinate axis; orthogonal box contact is exact.
    axis = max(range(3), key=lambda i: max(p[i] for p in new.values())
               - min(p[i] for p in new.values()))
    active: list[str] = []
    checks = 0
    for pid in sorted(boxes, key=lambda p: boxes[p][0][axis]):
        box = boxes[pid]
        active = [q for q in active if boxes[q][1][axis] + TOL_M >= box[0][axis]]
        for other in active:
            checks += 1
            if checks > MAX_CONTACT_CHECKS:
                report.errors.append("편집용 겹침 검사 범위를 초과했습니다. 최불리망 등 작은 망으로 나누어 출력하세요.")
                return
            if _touch(box, boxes[other]) and not _touch(old_boxes[pid], old_boxes[other]):
                report.errors.append(f"배관 {other} / {pid}: 격자 보정으로 새 겹침·접촉이 생깁니다.")
                return
        active.append(pid)


def prepare_kfp(data: dict, *, mode: str = "precise") -> tuple[dict | None, EditReport]:
    """Copy and audit KFP; return no editing copy when geometry is unsafe.

    ``precise`` preserves every existing calculation field. ``solver-edit``
    is an explicitly chosen, recalculation-required 1 cm copy. Its coordinates
    remain editable even when a legacy Solver drops grid metadata on saving.
    """
    if mode not in ("precise", "solver-edit"):
        raise ValueError("지원하지 않는 KFP 출력 모드입니다.")
    report = EditReport(mode)
    coords = _positions(data, report)
    pipes = data.get("pipe_data")
    if not isinstance(pipes, dict):
        report.errors.append("KFP 배관 정보가 없습니다.")
        pipes = {}
    report.nodes, report.pipes = len(coords), len(pipes)
    report.off_grid_nodes = [n for n, p in coords.items() if any(
        abs(v - round(v / SOLVER_GRID_M) * SOLVER_GRID_M) > TOL_M for v in p)]
    edges = {}
    for pid, pipe in pipes.items():
        if not isinstance(pipe, dict):
            report.errors.append(f"배관 {pid}: 배관 정보가 올바르지 않습니다.")
            continue
        a, b = pipe.get("start"), pipe.get("end")
        if a not in coords or b not in coords:
            report.errors.append(f"배관 {pid}: 연결 노드가 없습니다.")
            continue
        edges[pid] = (a, b)
        if math.dist(coords[a], coords[b]) <= TOL_M:
            report.errors.append(f"배관 {pid}: 길이가 0인 좌표입니다.")
        elif _axis(coords[a], coords[b]) is None:
            report.non_axis_pipes.append(str(pid))
    if report.errors:
        return None, report
    result = copy.deepcopy(data)
    if mode == "precise":
        result["grid_step"] = PRECISE_GRID_M if report.off_grid_nodes else SOLVER_GRID_M
        if report.off_grid_nodes:
            report.warnings.append(
                f"{len(report.off_grid_nodes)}개 노드에 미세 편집 격자를 지정했습니다. "
                "좌표·길이는 그대로입니다. Solver 저장 후 설정이 사라지면 오류가 재발할 수 있으므로 "
                "계속 편집하려면 ‘편집용 KFP’를 사용하세요.")
        if report.non_axis_pipes:
            report.warnings.append(f"축에 평행하지 않은 배관 {len(report.non_axis_pipes)}개는 원래 좌표를 보존했습니다.")
        return result, report

    if report.non_axis_pipes:
        report.errors.append("사선 배관은 자동 격자 보정하지 않습니다: " + ", ".join(report.non_axis_pipes[:12]))
        return None, report
    snapped = {n: tuple(round(v / SOLVER_GRID_M) * SOLVER_GRID_M for v in p)
               for n, p in coords.items()}
    occupied = {}
    for n, p in snapped.items():
        if p in occupied and math.dist(coords[n], coords[occupied[p]]) > TOL_M:
            report.errors.append(f"노드 {occupied[p]} / {n}: 격자 보정으로 서로 겹칩니다.")
        occupied[p] = n
        delta = max(abs(a - b) for a, b in zip(p, coords[n]))
        report.changed_nodes += int(delta > 1e-12)
        report.max_coordinate_change_mm = max(report.max_coordinate_change_mm, delta * 1000)
        meta = (result.get("nodes_meta_runtime") or {}).get(n)
        if isinstance(meta, dict):
            if "elevation_m" in meta:
                try:
                    elev = float(meta["elevation_m"])
                    if not math.isfinite(elev) or abs(elev - coords[n][2]) > TOL_M:
                        raise ValueError("schematic elevation")
                    meta["elevation_m"] = p[2]
                except (TypeError, ValueError):
                    report.errors.append(f"노드 {n}: 좌표 Z와 표고가 달라 편집용 보정을 중단합니다.")
            meta["coords"] = list(p)
        if isinstance(result.get("nodes"), dict) and n in result["nodes"]:
            result["nodes"][n] = list(p)
    for pid, (a, b) in edges.items():
        old_length = math.dist(coords[a], coords[b])
        new_length = math.dist(snapped[a], snapped[b])
        pipe = result["pipe_data"][pid]
        try:
            declared = float(pipe["length_m"])
            if not math.isfinite(declared) or abs(declared - old_length) > 1e-4:
                raise ValueError("schematic geometry")
        except (KeyError, TypeError, ValueError):
            report.errors.append(f"배관 {pid}: 좌표 거리와 계산 길이가 달라 편집용 보정을 중단합니다.")
            continue
        if new_length <= TOL_M or _axis(snapped[a], snapped[b]) != _axis(coords[a], coords[b]):
            report.errors.append(f"배관 {pid}: 격자 보정으로 배관이 소실되거나 방향이 바뀝니다.")
        delta = new_length - declared
        if abs(delta) > 1e-12:
            report.length_changes.append({"pipe": str(pid), "before_m": declared,
                                          "after_m": new_length, "delta_mm": delta * 1000})
        pipe["length_m"] = new_length
        # Existing hydraulic results must not masquerade as results for new lengths.
        for key in ("flow_lpm", "velocity_mps", "headloss_m"):
            if key in pipe:
                pipe[key] = 0.0
    if not report.errors:
        _new_contacts(coords, snapped, edges, report)
    report.max_length_change_mm = max((abs(c["delta_mm"]) for c in report.length_changes), default=0.0)
    report.total_length_change_mm = sum(c["delta_mm"] for c in report.length_changes)
    report.warnings.append("1cm 격자 편집용 사본입니다. 좌표·길이·표고가 바뀔 수 있으며 Solver에서 수리계산을 다시 실행해야 합니다.")
    result["grid_step"] = SOLVER_GRID_M
    result["module_f_editing_copy"] = report.to_dict()
    return (None if report.errors else result), report

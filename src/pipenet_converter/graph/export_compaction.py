"""Loss-preserving serial-pipe reduction on an export copy, never on the editor.

XY is CAD mm, elevation is m, declared pipe length is m. Only straight degree-2
points without a hydraulic attachment can disappear. Bends, junctions, heads,
boundaries, controls and review spots survive, including in cyclic networks.
"""
from __future__ import annotations

from collections import deque
from copy import deepcopy
from dataclasses import dataclass
import math
from typing import Iterable, Protocol, Sequence, TypeVar


class NetworkTables(Protocol):
    """Table-shaped intermediate representation shared by Module F writers."""

    nodes: list[dict]
    pipes: list[dict]
    nozzles: list[dict]
    fittings: list[dict]
    equipment: list[dict]


T = TypeVar('T', bound=NetworkTables)


@dataclass(frozen=True)
class CompactionAudit:
    """Export labels retain a reversible record of their original pipe rows."""

    nodes_before: int
    nodes_after: int
    pipes_before: int
    pipes_after: int
    removed_nodes: tuple[str, ...]
    source_pipes: dict[str, list[dict]]
    total_length_m: float

    @property
    def message(self) -> str:
        """Concise user-visible result without implying hydraulic approval."""
        return (f'출력용 직선 구간 정리: 노드 {self.nodes_before} → {self.nodes_after}, '
                f'배관 {self.pipes_before} → {self.pipes_after}. '
                '분기·헤드·부속·길이·루프 연결 보존, 편집 원본 변경 없음.')


# These fields describe provenance/display, not resistance. Every other field
# must agree; unfamiliar physical fields therefore fail conservatively closed.
_DIAGNOSTIC = {
    'label', 'in', 'out', 'length', 'elev', 'eq_len', 'bore_provenance',
    'dia_src', 'dia_source', 'dia_raw', 'dia_match_dist_mm', 'elev_source',
    'inferred_elev', 'inferred_length', 'off_tree',
}


def _ends(row: dict) -> tuple[str, str]:
    return str(row.get('in')), str(row.get('out'))


def protected_nodes(tables: NetworkTables) -> set[str]:
    """Find hydraulic attachment/review points before any loss aggregation."""
    nodes = {str(n['label']): n for n in tables.nodes}
    pipes = {str(p['label']): p for p in tables.pipes}
    protected = {k for k, n in nodes.items() if
        str(n.get('io_node', 'No')).lower() not in ('no', '', 'none')
        or n.get('pressure_pa') is not None or n.get('flow_m3s') is not None
        or n.get('protected') or n.get('head') or n.get('merge_reason')
        or n.get('terminal_pipe') or n.get('keep_node')}
    for name in ('nozzles', 'pumps', 'valves'):
        for row in getattr(tables, name, ()):
            protected.update(set(_ends(row)) & nodes.keys())
    heads = {str(h.get('in')) for h in tables.nozzles}
    for pipe in tables.pipes:
        if heads.intersection(_ends(pipe)):
            protected.update(_ends(pipe))  # Keep the actual head attachment/stem.
    for name in ('fittings', 'equipment'):
        for row in getattr(tables, name, ()):
            pipe = pipes.get(str(row.get('pipe')))
            node = str(row.get('editor_node') or row.get('node') or '')
            if pipe and node in _ends(pipe):
                protected.add(node)
            elif pipe:
                # Unknown/internal placement: do not alter the containing span.
                protected.update(_ends(pipe))
    unresolved = getattr(tables, 'unresolved', {}) or {}
    for rows in unresolved.values():
        if not isinstance(rows, list):
            continue
        for row in rows:
            if not isinstance(row, dict):
                continue
            pipe = pipes.get(str(row.get('pipe_label', row.get('pipe'))))
            node = str(row.get('node_label', row.get('node', '')))
            if node in nodes:
                protected.add(node)
            elif pipe:
                protected.update(_ends(pipe))
    return protected


def _straight(a: tuple, mid: tuple, b: tuple) -> bool:
    u = tuple(m-x for x, m in zip(a, mid))
    v = tuple(y-m for m, y in zip(mid, b))
    lu, lv = math.hypot(*u), math.hypot(*v)
    if min(lu, lv) < 1e-9:
        return False  # Coincident points may encode an unresolved vertical link.
    cosine = sum(x*y for x, y in zip(u, v)) / (lu*lv)
    return cosine >= 1-1e-9


def _unassigned_drawing_pipe(row: dict) -> bool:
    """An absent H diameter is not a conflicting or guessed diameter.

    Only a plain unmapped source span may be reduced. Generated head/riser
    segments and conflicting/partial matches remain explicit review boundaries.
    """
    rec = row.get('bore_provenance') or {}
    return (rec.get('policy') == 'drawing_first_v1'
            and rec.get('source') == 'unresolved' and rec.get('block_export') is True
            and row.get('dia') in (None, 0) and row.get('inner_mm') in (None, 0)
            and rec.get('reason') == 'no_match'
            and not any(rec.get(key) for key in ('text_mm', 'candidate_mm', 'manual',
                                                 'reference_annotation', 'review_reasons')))


def _compatible(a: dict, b: dict, *, allow_unassigned: bool = False) -> bool:
    unassigned = allow_unassigned and all(_unassigned_drawing_pipe(p) for p in (a, b))
    if not unassigned and any((p.get('bore_provenance') or {}).get('block_export') for p in (a, b)):
        return False
    if any(p.get('waypoints') or str(p.get('status', 'Normal')).lower() != 'normal'
           for p in (a, b)):
        return False
    if any((p.get('dia') is None and not unassigned) or p.get('c') is None or not p.get('type') for p in (a, b)):
        return False
    if any(not math.isfinite(float(p.get('length', 0))) or float(p.get('length', 0)) <= 0
           or p.get('eq_len', 0) is None or not math.isfinite(float(p.get('eq_len', 0)))
           or float(p.get('eq_len', 0)) < 0 for p in (a, b)):
        return False
    left = {k: v for k, v in a.items() if k not in _DIAGNOSTIC}
    right = {k: v for k, v in b.items() if k not in _DIAGNOSTIC}
    return left == right


def compact_export(tables: T, *, keep_nodes: Iterable[str] = (),
                   display_views: Sequence[Sequence[dict]] = (),
                   allow_unassigned: bool = False) -> tuple[T, CompactionAudit]:
    """Return a compact COPY and its audit, preserving every physical attachment.

    Only consistently directed serial pipes are combined: opposing solver
    reference directions remain explicit. Extra display views must also remain
    straight so schematic corners never disappear. Lengths are summed, not
    remeasured from drawing coordinates. Labels of retained nodes never change.
    H canonicalization may opt into joining plain unassigned source spans;
    this never supplies a diameter or authorizes calculation-ready export.
    """
    result = deepcopy(tables)
    nodes = {str(n['label']): n for n in result.nodes}
    pipes = {str(p['label']): p for p in result.pipes}
    if len(nodes) != len(result.nodes) or len(pipes) != len(result.pipes):
        raise ValueError('출력 정리: 중복 노드/배관 이름이 있습니다.')
    adjacent: dict[str, set[str]] = {k: set() for k in nodes}
    for key, pipe in pipes.items():
        for end in _ends(pipe):
            if end not in nodes:
                raise ValueError(f'출력 정리: 배관 {key}의 연결 노드 {end}가 없습니다.')
            adjacent[end].add(key)
    protected = protected_nodes(result) | set(map(str, keep_nodes))
    xyz = {k: (float(n['x'])/1000, float(n['y'])/1000, float(n.get('elevation', 0)))
           for k, n in nodes.items()}
    views = [{str(n['label']): (float(n['x']), float(n['y'])) for n in view}
             for view in display_views if view]
    if any(not math.isfinite(v) for pt in xyz.values() for v in pt):
        raise ValueError('출력 정리: 유한하지 않은 좌표가 있습니다.')
    sources = {k: [deepcopy(p)] for k, p in pipes.items()}
    order = {k: i for i, k in enumerate(pipes)}
    removed = []
    queue = deque(nodes)
    while queue:
        node = queue.popleft()
        if node not in nodes or node in protected or len(adjacent[node]) != 2:
            continue
        pa, pb = sorted(adjacent[node], key=order.__getitem__)
        a, b = pipes[pa], pipes[pb]
        if str(a['out']) != node or str(b['in']) != node:
            pa, pb, a, b = pb, pa, b, a
        if str(a['out']) != node or str(b['in']) != node or not _compatible(a, b, allow_unassigned=allow_unassigned):
            continue
        start, end = str(a['in']), str(b['out'])
        if start == end or start == node or end == node:
            continue
        if not _straight(xyz[start], xyz[node], xyz[end]):
            continue
        if any(not all(k in view for k in (start, node, end))
               or not (_straight(view[start], view[node], view[end])
                       or view[start] == view[node] == view[end]) for view in views):
            continue
        # Do not collapse a cycle into a parallel/single edge unsupported by a
        # downstream consumer. No spanning-tree pruning is ever performed.
        if (adjacent[start] & adjacent[end]) - {pa, pb}:
            continue
        # Elevation declarations must agree with node values before reduction.
        if any(abs(float(p.get('elev', xyz[str(p['out'])][2]-xyz[str(p['in'])][2]))
                   - (xyz[str(p['out'])][2]-xyz[str(p['in'])][2])) > .001 for p in (a, b)):
            continue
        label = min((pa, pb), key=order.__getitem__)
        merged = deepcopy(a)
        merged.update(label=label, length=float(a['length'])+float(b['length']),
                      elev=xyz[end][2]-xyz[start][2],
                      eq_len=float(a.get('eq_len', 0))+float(b.get('eq_len', 0)))
        merged['out'] = b['out']
        # Fittings are at protected endpoints. Update their host pipe, without
        # removing/duplicating them or their directional loss assignment.
        for row in result.fittings + result.equipment:
            host = str(row.get('pipe'))
            if host not in (pa, pb):
                continue
            if 'rel_pos' in row:
                offset = 0 if host == pa else float(a['length'])
                old_length = float(pipes[host]['length'])
                row['rel_pos'] = (offset + float(row['rel_pos'])*old_length)/merged['length']
            row['pipe'] = label
            for key in ('in', 'out'):
                if key in row:
                    row[key] = merged[key]
        history = sources.pop(pa)
        history.extend(sources.pop(pb))
        for pid in (pa, pb):
            for endpoint in _ends(pipes.pop(pid)):
                adjacent[endpoint].discard(pid)
        sources[label] = history
        pipes[label] = merged
        adjacent[start].add(label)
        adjacent[end].add(label)
        del nodes[node]
        removed.append(node)
        queue.extend((start, end))
    result.nodes = [n for n in result.nodes if str(n['label']) in nodes]
    result.pipes = [pipes[k] for k in sorted(pipes, key=order.__getitem__)]
    # Construct provenance only once per final span, not once per removed node
    # (which would repeatedly copy a growing history on long straight runs).
    for pipe in result.pipes:
        originals = sources[str(pipe['label'])]
        if len(originals) > 1 and any('bore_provenance' in p for p in originals):
            pipe['bore_provenance'] = dict(method='export_serial_merge',
                block_export=any((p.get('bore_provenance') or {}).get('block_export') for p in originals),
                review_only=any(p.get('review_only') or (p.get('bore_provenance') or {}).get('review_only')
                                for p in originals),
                source_pipes=[str(p['label']) for p in originals])
    # Editor-to-table label maps are provenance too; do not leave dangling IDs.
    rename = {str(p['label']): key for key, rows in sources.items() for p in rows}
    if hasattr(result, 'pipe_labels'):
        result.pipe_labels = {k: rename.get(str(v), v) for k, v in result.pipe_labels.items()}
    if hasattr(result, 'node_labels'):
        result.node_labels = {k: v for k, v in result.node_labels.items() if str(v) in nodes}
    # Preserve review spot host references (their protected nodes still exist).
    for rows in (getattr(result, 'unresolved', {}) or {}).values():
        if isinstance(rows, list):
            for row in rows:
                if isinstance(row, dict):
                    for key in ('pipe_label', 'pipe'):
                        if key in row and str(row[key]) in rename:
                            row[key] = rename[str(row[key])]
    total = math.fsum(float(p['length']) for p in tables.pipes)
    if not math.isclose(total, math.fsum(float(p['length']) for p in result.pipes), abs_tol=1e-8):
        raise ValueError('출력 정리 전후 배관 연장이 다릅니다.')
    original_labels = [str(p['label']) for rows in sources.values() for p in rows]
    if (len(original_labels) != len(set(original_labels))
            or set(original_labels) != {str(p['label']) for p in tables.pipes}
            or set(sources) != {str(p['label']) for p in result.pipes}):
        raise ValueError('연속관 대응 오류: 원본 배관의 중복 또는 누락이 있습니다.')
    return result, CompactionAudit(len(tables.nodes), len(result.nodes),
        len(tables.pipes), len(result.pipes), tuple(removed), sources, total)

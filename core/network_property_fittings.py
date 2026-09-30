"""Explicit edits of existing fitting rows, retaining positions and hydraulic totals."""
from __future__ import annotations

from core.network_editor import EditError, Network, fitting_value, number


def fitting_rows(network: Network, catalog: dict) -> list[dict]:
    """Expose pipe-local identities, stable across merged pipe label translation."""
    result = []
    aliases = catalog.get('fitting_aliases', {})
    names = {r['id']: r['name'] for r in catalog['fittings']}
    for collection in ('fittings', 'equipment'):
        indices: dict[str, int] = {}
        for row in getattr(network.tables, collection):
            pipe = str(row.get('pipe'))
            index = indices.get(pipe, 0)
            indices[pipe] = index + 1
            kind = row.get('type') if collection == 'fittings' else row.get('editor_library')
            library = aliases.get(kind, kind)
            # Pumps/valves and unknown devices must not become an elbow implicitly.
            if not library or library not in names:
                continue
            p = network.pipes.get(pipe)
            if p is None:
                continue
            node = row.get('node') if collection == 'fittings' else row.get('editor_node')
            if collection == 'fittings' and node is None:
                node = p.a
            result.append(dict(label=f'{collection}:{pipe}:{index}', pipe=pipe,
                collection=collection, index=index, node=str(node) if node is not None else None,
                position=row.get('rel_pos', .5), library=library, name=names[library],
                count=row.get('count', 1), expected_type=kind))
    return result


def update_fitting(network: Network, command: dict, catalog: dict) -> None:
    """Change an existing library fitting, never append or move a loss element."""
    target = str(command['target'])
    row_info = next((r for r in fitting_rows(network, catalog)
                    if r['pipe'] == target and r['collection'] == command.get('collection')
                    and r['index'] == command.get('index')), None)
    if row_info is None or row_info['expected_type'] != command.get('expected_type'):
        raise EditError('부속 원본이 변경되었습니다. 다시 선택하세요.')
    rows = [r for r in getattr(network.tables, row_info['collection']) if str(r.get('pipe')) == target]
    row = rows[row_info['index']]
    count_value = number(command.get('count', row_info['count']), '부속 개수', 1, 100)
    if int(count_value) != count_value:
        raise EditError('부속 개수는 정수여야 합니다.')
    count = int(count_value)
    library = str(command.get('fitting') or row_info['library'])
    dn = int(network.pipes[target].row.get('dia') or 0)
    value = fitting_value(catalog, library, dn) * count
    if row_info['collection'] == 'fittings':
        old = fitting_value(catalog, row_info['library'], dn) * int(row_info['count'])
        total = network.pipes[target].row.get('eq_len')
        if total is None or float(total) + value - old < -1e-7:
            raise EditError('저장된 부속 손실 합계를 먼저 확인하세요.')
        aliases = catalog.get('fitting_aliases', {})
        kind = next((key for key, lib in aliases.items() if lib == library), library)
        row.update(type=kind, count=count)
        network.pipes[target].row['eq_len'] = max(0., float(total) + value - old)
    else:
        name = next(r['name'] for r in catalog['fittings'] if r['id'] == library)
        row.update(editor_library=library, count=count, eq_len=value, desc=name)
    row['editor_note'] = str(command.get('note') or 'H 부속 속성 수정')

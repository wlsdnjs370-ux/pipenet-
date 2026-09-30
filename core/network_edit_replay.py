"""Replay edits over a changed source graph without aliasing generated labels."""
from __future__ import annotations

from copy import deepcopy

from core.network_editor import EditError, Network, apply_edit

NODE_OPS = {'extend', 'connect', 'head', 'node', 'move_node', 'delete_node', 'merge_node', 'paste'}


def rebase_commands(old: Network, new: Network, commands: list[dict], cursor: int,
                    catalog: dict) -> tuple[Network, list[dict]]:
    """Translate generated identities while replaying, preserving physical offsets.

    An active edit whose source disappeared rejects the entire transaction. Redo
    commands may refer to a source restored by a later redo in the other scope;
    retain those commands until that source exists again.
    """
    def collections(net: Network) -> dict:
        return {'node': net.nodes, 'pipe': net.pipes,
                'equipment': {str(r['label']): r for r in net.tables.equipment}}

    old_sets, new_sets = collections(old), collections(new)
    maps = {kind: {label: label for label in rows if label in new_sets[kind]}
            for kind, rows in old_sets.items()}
    accepted = deepcopy(new)
    translated = []
    for index, source in enumerate(commands):
        command = deepcopy(source)
        if command.get('op') == 'batch_properties':
            # Property-only batches introduce no graph identities. Translate each
            # child with the same mapping and keep the entire history unit atomic.
            children = command['commands']
            try:
                for child in children:
                    namespace = 'node' if child.get('op') in NODE_OPS else 'pipe'
                    child['target'] = maps[namespace][str(child['target'])]
                old_next, _ = apply_edit(old, source, catalog)
                new_next, _ = apply_edit(new, command, catalog)
                old, new = old_next, new_next
                translated.append(command)
                if index < cursor:
                    accepted = new
                continue
            except (EditError, KeyError) as exc:
                if index < cursor:
                    raise EditError(f'통합 일괄 편집 {index+1}번을 유지할 수 없습니다: {exc}') from exc
                translated.extend(deepcopy(commands[index:]))
                break
        kind = 'node' if command.get('op') in NODE_OPS else 'pipe'
        try:
            fields = [] if command['op']=='compact_runs' else [('target',kind)]
            if command['op']=='compact_runs':
                command['keep'] = [maps['node'].get(str(k),str(k)) for k in command.get('keep',())]
            if command['op']=='connect': fields.append(('end','node'))
            if command['op']=='paste': fields.extend((('source','node'),('source_pipe','pipe')))
            if command['op']=='fitting' and command.get('node'): fields.append(('node','node'))
            if command['op']=='remove_fitting': fields.append(('equipment','equipment'))
            for field, namespace in fields:
                if field not in command:
                    continue
                label = str(command[field])
                if label not in maps[namespace]:
                    raise EditError(f'{namespace} {label}: 원본 요소가 없어 통합 편집을 옮길 수 없습니다.')
                command[field] = maps[namespace][label]
            if command['op'] == 'move_node':
                before = old.nodes[str(source['target'])].xyz
                after = new.nodes[command['target']].xyz
                for i, field in enumerate(('x', 'y', 'z')):
                    command[field] = float(source[field]) - before[i] + after[i]
            old_next, _ = apply_edit(old, source, catalog)
            new_next, _ = apply_edit(new, command, catalog)
            for namespace, rows in collections(old_next).items():
                added_old = [k for k in rows if k not in collections(old)[namespace]]
                added_new = [k for k in collections(new_next)[namespace] if k not in collections(new)[namespace]]
                if len(added_old) != len(added_new):
                    raise EditError('통합 편집의 생성 요소 수가 달라졌습니다.')
                maps[namespace].update(zip(added_old, added_new))
            old, new = old_next, new_next
            translated.append(command)
            if index < cursor:
                accepted = new
        except (EditError, KeyError) as exc:
            if index < cursor:
                raise EditError(f'통합 편집 {index+1}번을 유지할 수 없습니다: {exc}') from exc
            translated.extend(deepcopy(commands[index:]))
            break
    return accepted, translated

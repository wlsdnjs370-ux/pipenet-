"""Shared live graph editing endpoints for calculation and merged networks."""
from __future__ import annotations

from copy import deepcopy
import functools
import hashlib
import json
from dataclasses import asdict

from flask import jsonify, request

from core.network_editor import EditError, apply_edit
from core.network_edit_replay import NODE_OPS, rebase_commands
from routes.module_f import network_edit as ne
from routes.module_f import editor_history
from routes.module_f.common import _fail
from routes.module_f.jobs import _job_running, route_session


def _target_scope(sess: dict, scope: str, command: dict) -> tuple[str, dict, int]:
    """Plan objects in the merged view edit the original plan, then re-merge."""
    if command.get('op') == 'batch_properties' and scope == 'merge':
        children = command.get('commands')
        if not isinstance(children, list) or not children or any(not isinstance(c, dict) or c.get('op') == 'batch_properties' for c in children):
            raise EditError('일괄 수정 대상을 확인하세요.')
        mapped = [_target_scope(sess, scope, c) for c in children]
        if len({(s, offset) for s, _, offset in mapped}) != 1:
            raise EditError('평면도와 계통도·기계실 대상은 나누어 수정하세요. 변경은 저장되지 않았습니다.')
        return mapped[0][0], dict(command, commands=[c for _, c, _ in mapped]), mapped[0][2]
    if scope != 'merge':
        return scope,command,0
    from routes.module_f.merge import label_offset_for
    obj = sess['merged']
    nodes = set(map(str,(obj.get('parts') or {}).get('plan',())))
    target = str(command.get('target'))
    kind = command.get('op')
    if kind == 'compact_runs':
        # A merged plan is not a second editable copy. Normalize it at its
        # owner first and rebuild the integrated view through the same history.
        plan = ne.ensure(sess,'design')
        _, preview = apply_edit(plan['current'],ne.compaction_command(sess,'design'),ne.catalog())
        if preview['compaction']['removed_nodes']:
            return 'design',command,0
        return scope,command,0
    is_node = kind in NODE_OPS
    row = next((r for r in obj['combined'].pipes if str(r['label'])==target),None)
    in_plan = target in nodes if is_node else bool(row and str(row['in']) in nodes and str(row['out']) in nodes)
    source_target = (obj.get('plan_pipe_sources') or {}).get(target, target)
    if kind == 'connect' and str(command.get('end')) not in nodes:
        in_plan = False
    if kind == 'paste' and str(command.get('source')) not in nodes:
        in_plan = False
    if not in_plan:
        return scope,command,0
    design = ne.plan_state(sess).get('design') or {}
    offset = label_offset_for(design.get('method','manual'))
    original_net = ne.ensure(sess,'design')['current']
    # A new merge-only branch can inherit part=plan, but has no source in 04.
    if is_node:
        if not target.isdigit() or str(int(target)-offset) not in original_net.nodes:
            return scope,command,0
    elif source_target not in original_net.pipes:
        return scope,command,0
    for field in (('end',) if kind=='connect' else ('source',) if kind=='paste' else ()):
        label=str(command.get(field,''))
        if not label.isdigit() or str(int(label)-offset) not in original_net.nodes:
            return scope,command,0
    command = deepcopy(command)
    if not is_node:
        command['target'] = source_target
    if is_node:
        try:
            command['target'] = str(int(target)-offset)
            if kind == 'move_node':
                from core.network_editor import number
                shown = ne.ensure(sess,'merge')['current'].nodes[target].xyz
                original = ne.ensure(sess,'design')['current'].nodes[command['target']].xyz
                for i,key in enumerate(('x','y','z')):
                    command[key] = number(command.get(key),key,-1e7,1e7)-shown[i]+original[i]
            if kind == 'connect':
                command['end'] = str(int(command['end'])-offset)
            if kind == 'paste':
                command['source'] = str(int(command['source'])-offset)
        except ValueError:
            raise EditError('이음매는 통합에서 별도 편집해야 합니다.') from None
    if command.get('node'):
        command['node'] = str(int(command['node'])-offset)
    return 'design',command,offset


RISER_OPS = ('delete_join', 'delete_cut')


def _merged_selection(selection: dict, merged: dict, offset: int) -> dict:
    """Translate an edited source selection back to its displayed merged identity."""
    if selection['kind'] == 'node':
        return dict(selection, label=str(int(selection['label']) + offset))
    sources = merged.get('plan_pipe_sources') or {}
    label = next((shown for shown, original in sources.items()
                  if original == str(selection['label'])), selection['label'])
    return dict(selection, label=label)


def _riser_command(sess: dict, scope: str, command: dict) -> dict:
    """[오너 2026-09-21] 세로관 삭제 카드(예/아니오)는 통합망의 계통도 배관에만 쓴다.

    예(delete_join)가 붙일 때 남겨야 할 노드 — 기준점 10 과 기계실이 붙는
    노드 — 를 명령에 적어 둔다. 편집 기록을 다시 돌려도 같은 노드가 남는다.
    """
    if command.get('op') not in RISER_OPS:
        return command
    if scope != 'merge':
        raise EditError('세로관 삭제(위·아래 연결)는 통합 배관망의 계통도 배관에서만 할 수 있습니다.')
    from routes.module_f.merge import ANCHOR_LABEL
    got = sess['merged']
    target = str(command.get('target'))
    part = (got.get('pipe_parts') or {}).get(target)
    if part is None:
        system = {str(x) for x in (got.get('parts') or {}).get('system', ())}
        row = next((r for r in got['combined'].pipes if str(r['label']) == target), None)
        part = 'system' if row and {str(row['in']), str(row['out'])} <= system else None
    if part != 'system':
        raise EditError('계통도 세로관만 이 방법으로 지울 수 있습니다.')
    keep = {ANCHOR_LABEL} | ({str(got['pump_junction'])} if got.get('pump_junction') else set())
    return dict(command, keep=sorted(keep))


def _revision(editor: dict) -> str:
    return hashlib.sha256(json.dumps(dict(graph=ne.fingerprint(editor['current']),
                          commands=editor['commands'],cursor=editor['cursor']),
                          sort_keys=True).encode()).hexdigest()


def _rebuild_merge(sess: dict) -> None:
    if not (sess.get('merged') or {}).get('combined'):
        return
    from routes.module_f.api_merge import rebuild_merged
    previous = ne.ensure(sess,'merge')
    if previous['conflict']:
        raise EditError(previous['conflict'])
    rebuild_merged(sess, persist_overrides=False)
    obj = sess['merged']
    base = ne._network(sess,'merge',obj)
    current,commands = rebase_commands(previous['base'],base,previous['commands'],
                                       previous['cursor'],ne.catalog())
    fresh = dict(previous,object=obj,base_object=deepcopy(obj),base=base,
                 base_hash=ne.fingerprint(base),current=current,commands=commands,conflict=None)
    # Rebased commands/geometry cannot reuse checkpoints from the old merge.
    fresh.pop('_replay_cache', None)
    fresh['revision'] = _revision(fresh)
    ne.sync_object(sess,'merge',obj,current)
    sess['merge_editor'] = fresh


def _order(command: dict, scope: str, index: int) -> int:
    # Legacy histories could only contain plan edits followed by merge edits.
    return command.get('_order',index + (1000000 if scope=='merge' else 0))


def _history_scope(sess: dict, action: str) -> str:
    choices = []
    for scope in ('design','merge'):
        editor = ne.ensure(sess,scope)
        i = editor['cursor']-1 if action=='undo' else editor['cursor']
        if 0 <= i < len(editor['commands']):
            choices.append((_order(editor['commands'][i],scope,i),scope))
    if not choices:
        return 'merge'
    return (max(choices) if action=='undo' else min(choices))[1]


def _next_order(sess: dict) -> int:
    values = [1000000]
    for scope in ('design','merge'):
        if scope=='merge' and not (sess.get('merged') or {}).get('combined'):
            continue
        e = ne.ensure(sess,scope)
        values.extend(_order(c,scope,i) for i,c in enumerate(e['commands']))
    return max(values)+1


def commit_edit(sess: dict, scope: str, editor: dict, candidate, commands: list,
                cursor: int, *, branch: bool = False, propagate: bool = True) -> None:
    """Persist and accept together. Protected from cancellation mid-commit."""
    fresh = dict(editor,current=candidate,commands=commands,cursor=cursor,conflict=None)
    editor_history.remember(fresh)
    fresh['revision'] = _revision(fresh)
    # Work on a copy first: a failed merge cannot leave the plan half edited.
    state,key,obj = ne._container(sess,scope)
    old_obj = deepcopy(obj)
    old_merged = deepcopy(sess.get('merged'))
    old_merge_editor = sess.get('merge_editor')
    old_plan_editor = ne.plan_state(sess).get('network_editor')
    old_summary = sess.get('merge_summary')
    plan = ne.plan_state(sess)
    absent = object()
    restore_plan = {k:plan.get(k,absent) for k in
                    ('design_sdf_path','design_slf_path','design_has_path','worst_kfp_path')}
    restore_merge = {k:sess.get(k,absent) for k in ('merge_files','merge_missed')}
    old_auto_tables = (plan.get('auto') or {}).get('tables',absent)
    file_before, written = {}, {}
    try:
        # New edits discard the undone branch in both views, as one timeline.
        if branch:
            other_scope = 'merge' if scope=='design' else 'design'
            if other_scope=='design' or (sess.get('merged') or {}).get('combined'):
                other = ne.ensure(sess,other_scope)
                other_state,other_key,_ = ne._container(sess,other_scope)
                other = dict(other,commands=other['commands'][:other['cursor']])
                editor_history.remember(other)
                other['revision'] = _revision(other)
                other_state[other_key] = other
        obj.clear()
        obj.update(deepcopy(editor['base_object']))
        ne.sync_object(sess,scope,obj,candidate)
        if scope == 'design' and propagate:
            _rebuild_merge(sess)
        fresh['revision'] = _revision(fresh)
        pending = {scope:fresh}
        if propagate and (sess.get('merged') or {}).get('combined'):
            other_scope = 'merge' if scope=='design' else 'design'
            pending[other_scope] = dict(ne.ensure(sess,other_scope))
        for save_scope,e in pending.items():
            path = ne._path(sess,save_scope)
            file_before[path] = path.read_bytes() if path.is_file() else None
            ne.save_history(sess,save_scope,e)
            written[path] = e['disk_hash']
            other_state,other_key,_ = ne._container(sess,save_scope)
            other_state[other_key] = e
    except Exception:
        for path, digest in written.items():
            if path.is_file() and hashlib.sha256(path.read_bytes()).hexdigest()==digest:
                if file_before[path] is None:
                    path.unlink()
                else:
                    tmp = path.with_suffix('.tmp')
                    tmp.write_bytes(file_before[path]); tmp.replace(path)
        obj.clear()
        obj.update(old_obj)
        sess['merged'] = old_merged
        if old_merge_editor and old_merged:
            old_merge_editor = dict(old_merge_editor,object=old_merged)
        sess['merge_editor'] = old_merge_editor
        ne.plan_state(sess)['network_editor'] = old_plan_editor
        sess['merge_summary'] = old_summary
        for container, fields in ((plan,restore_plan),(sess,restore_merge)):
            for field,value in fields.items():
                if value is absent: container.pop(field,None)
                else: container[field] = value
        if plan.get('auto') is not None:
            if old_auto_tables is absent: plan['auto'].pop('tables',None)
            else: plan['auto']['tables'] = old_auto_tables
        raise
    state[key] = fresh
    if scope == 'design':
        sess['editor_last_scope'] = 'design'
    else:
        sess['editor_last_scope'] = 'merge'


def register(app) -> None:
    """Register one editor API shared by both UI stages."""
    @app.get('/api/module-f/network-editor')
    @route_session()
    def module_f_network_editor_state(sess, body):
        try:
            scope = 'merge' if body.get('scope')=='merge' else 'design'
            with ne.EDIT_LOCK:
                if _job_running(sess):
                    return _fail('진행 중인 작업을 마친 뒤 편집하세요.',409)
                editor = ne.ensure(sess,scope)
                result = ne.public_state(editor)
                result['notice'] = editor.get('notice')
                basis_stale = None
                if scope == 'design' and editor['conflict']:
                    from routes.module_f.selection import _design_stale
                    plan = ne.plan_state(sess)
                    basis_stale = _design_stale(dict(plan, network_editor=dict(editor, conflict=None)))
                result['archived'] = ne.archived_histories(sess,scope) if body.get('history')=='1' else []
                # Plan undo is available in the integrated view as well.
                if scope=='merge':
                    plan = ne.ensure(sess,'design')
                    if body.get('history')=='1':
                        result['archived'] += ne.archived_histories(sess,'design')
                    result['undo'] += plan['cursor']
                    result['redo'] += len(plan['commands'])-plan['cursor']
                    result['history'] = sorted(plan['commands']+editor['commands'],key=lambda c:c.get('_order',0))
                    result['history_scope'] = _history_scope(sess,'undo')
                return jsonify(ok=True,sid=sess['id'],scope=scope,
                               can_accept_basis=bool(scope=='design' and editor['conflict'] and not basis_stale),
                               basis_stale=basis_stale,
                               catalog=ne.catalog(),**result)
        except (ValueError,OSError) as exc:
            return _fail(str(exc),409)

    @app.post('/api/module-f/network-editor')
    @route_session(post=True)
    def module_f_network_editor_apply(sess, body):
        try:
            scope = 'merge' if body.get('scope')=='merge' else 'design'
            requested_scope = scope
            with ne.EDIT_LOCK:
                if _job_running(sess):
                    raise EditError('진행 중인 작업을 마친 뒤 편집하세요.')
                shown = ne.ensure(sess,scope)
                if body.get('revision') != shown['revision']:
                    return _fail('배관망이 바뀌었습니다. 다시 선택하고 미리보기를 확인하세요.',409)
                action = body.get('action','preview')
                design = ne.plan_state(sess).get('design') or {}
                if action == 'accept_basis':
                    # This is a user-confirmed recovery, never an implicit reset
                    # on field entry/rebuild. Old labels cannot target a new graph.
                    if scope != 'design' or not shown['conflict']:
                        raise EditError('새 기준망으로 전환할 편집 충돌이 없습니다.')
                    from routes.module_f.selection import _design_stale
                    plan = ne.plan_state(sess)
                    probe = dict(plan, network_editor=dict(shown, conflict=None))
                    if _design_stale(probe):
                        raise EditError('도면 선정이 다시 바뀌었습니다. 입력값을 먼저 반영한 뒤 계속하세요.')
                    from routes.module_f.cancellation import checkpoint
                    checkpoint()
                    ne.backup_history(sess,scope,shown.get('disk_hash'))
                    # No rejected command is replayed. commit_edit is atomic and
                    # invalidates generated outputs. The old integrated view is
                    # not silently rebuilt from a newly accepted plan.
                    candidate = deepcopy(shown['base'])
                    commands = []
                    command = ne.compaction_command(sess, scope)
                    candidate, normalized = apply_edit(candidate, command, ne.catalog())
                    if normalized['compaction']['removed_nodes']:
                        command.update(_order=_next_order(sess), note='표 확정 후 연속 동일관 자동 정리')
                        commands.append(command)
                    commit_edit(sess,scope,shown,candidate,commands,len(commands),propagate=False)
                    fresh = ne.ensure(sess,scope)
                    fresh['notice'] = '이전 편집 기록을 별도 보관하고 새 기준 배관망으로 전환했습니다.'
                    return jsonify(ok=True,scope=scope,selection=None,notice=fresh['notice'])
                if action != 'reset' and design.get('sig') is not None:
                    from routes.module_f.selection import _design_stale
                    if _design_stale(ne.plan_state(sess)):
                        raise EditError('도면 손질·선정이 바뀌어 현재 표가 이전 상태입니다. 편집 기록을 확인하고 표를 다시 확정하세요.')
                if shown['conflict'] and action != 'reset':
                    raise EditError(shown['conflict'])
                if scope == 'merge':
                    sess['editor_merge_iso'] = bool(body.get('iso',False))
                    from core.network_editor import number
                    sess['editor_merge_z'] = number(body.get('iso_z_scale',1),'표시 배율',.1,10)
                command = dict(body.get('command') or {})
                if action in ('preview','apply') and command.get('op') == 'paste':
                    if command.get('copied_revision') != shown['revision']:
                        raise EditError('복사 이후 배관망이 바뀌었습니다. 원본 노드를 다시 복사하세요.')
                offset = 0
                if action in ('preview','apply'):
                    scope,command,offset = _target_scope(sess,scope,command)
                    if command.get('op') == 'pipe_reference':
                        # Resolve identity against this drawing, never trust client
                        # text/coordinates/DN. Store the validated snapshot for undo/replay.
                        plan = ne.plan_state(sess)
                        if scope != 'design' or (plan.get('design_settings') or {}).get('diameter_policy') != 'drawing_first_v1':
                            raise EditError('평면도 입력값 정의에서 원본 표기를 선택하세요.')
                        annotation = next((a for a in plan.get('_h_diameter_annotations', [])
                                           if a.identity == command.get('annotation_id') and a.raw_text.strip()), None)
                        if annotation is None:
                            raise EditError('현재 도면에 없는 원본 관경 표기입니다. 다시 선택하세요.')
                        command = dict(op='pipe_reference', target=str(command.get('target', '')),
                                       annotation_id=annotation.identity, annotation=asdict(annotation),
                                       note='원본 도면 참조 내경 변경')
                    command = _riser_command(sess,scope,command)
                    if command.get('op') == 'compact_runs':
                        command = ne.compaction_command(sess,scope,target=str(command.get('target','')))
                elif scope=='merge' and action in ('undo','redo'):
                    scope = _history_scope(sess,action)
                elif scope=='merge' and action=='reset' and not shown['commands']:
                    scope = 'design'
                if requested_scope=='merge' and scope=='design':
                    from routes.module_f.merge import label_offset_for
                    offset = label_offset_for(design.get('method','manual'))
                editor = ne.ensure(sess,scope)
                commands = deepcopy(editor['commands'])
                cursor = editor['cursor']
                selection = None
                if action in ('preview','apply'):
                    command['_order'] = _next_order(sess)
                    candidate,selection = apply_edit(editor['current'],command,ne.catalog())
                    commands = commands[:cursor]+[command]
                    cursor += 1
                elif action in ('undo','redo','reset'):
                    if action=='reset':
                        commands,cursor = [],0
                    else:
                        cursor = max(0,min(len(commands),cursor+(-1 if action=='undo' else 1)))
                    candidate,selection = editor_history.replay(
                        editor, commands, cursor, ne.catalog(), sid=sess.get('id',''))
                else:
                    raise EditError('지원하지 않는 편집 요청입니다.')
                if action=='preview':
                    view = ne.project(sess,scope,candidate)
                    if requested_scope=='merge' and scope=='design':
                        # Use an isolated re-merge, not the plan projection, in
                        # the merged canvas (its coordinate origin is different).
                        clone = dict(sess)
                        clone['element_overrides'] = deepcopy(sess.get('element_overrides'))
                        clone['element_ops'] = deepcopy(sess.get('element_ops'))
                        from routes.module_f.slots import _slot_active
                        if _slot_active(sess)=='plan':
                            clone['design'] = deepcopy(sess['design'])
                        else:
                            clone['slots'] = dict(sess['slots'])
                            clone['slots']['plan'] = dict(ne.plan_state(sess))
                            clone['slots']['plan']['design'] = deepcopy(ne.plan_state(sess)['design'])
                        if ne.plan_state(sess).get('auto'):
                            ne.plan_state(clone)['auto'] = dict(ne.plan_state(sess)['auto'])
                        ne.sync_object(clone,'design',ne.plan_state(clone)['design'],candidate)
                        _rebuild_merge(clone)
                        view = ne.project(clone,'merge',clone['merge_editor']['current'])
                        selection = _merged_selection(selection, clone['merged'], offset)
                        selection['counts'] = dict(nodes=len(view['nodes']),pipes=len(view['pipes']),
                            heads=len(clone['merged']['combined'].nozzles))
                    return jsonify(ok=True,preview=view,selection=selection,
                                   counts=selection['counts'],revision=shown['revision'],effective_scope=scope)
                from routes.module_f.cancellation import checkpoint
                checkpoint()
                commit_edit(sess,scope,editor,candidate,commands,cursor,branch=action=='apply')
                if (body.get('response_mode') == 'reference_patch' and action == 'apply'
                        and command.get('op') == 'pipe_reference' and requested_scope == scope == 'design'):
                    # Reference changes do not alter geometry. Return the committed
                    # row instead of requiring projection + full inspector reloads.
                    fresh = ne.plan_state(sess)['network_editor']
                    pipe = candidate.pipes[command['target']]
                    attribute_patch = None
                    plan = ne.plan_state(sess)
                    if plan.get('_h_diameter_context') and plan.get('edit'):
                        from routes.module_h_attributes import state as attribute_state
                        snapshot = attribute_state(plan)
                        attribute_patch = dict(revision=snapshot['revision'], pending=snapshot['pending'],
                            can_undo=snapshot['can_undo'], record=next(r for r in snapshot['records']
                                if r['kind'] == 'pipe' and r['label'] == command['target']))
                    return jsonify(ok=True, scope=scope, selection=selection,
                        reference_patch=dict(revision=fresh['revision'], undo=fresh['cursor'],
                            redo=len(fresh['commands'])-fresh['cursor'], history=fresh['commands'],
                            pipe=dict(pipe.row, **{'in':pipe.a, 'out':pipe.b}),
                            equipment=candidate.tables.equipment,
                            nodes=[dict(label=k,merge_reason=ne.merge_reason(candidate,k)) for k in (pipe.a,pipe.b)],
                            attributes=attribute_patch))
                if selection and requested_scope=='merge' and scope=='design':
                    selection = _merged_selection(selection, sess['merged'], offset)
                return jsonify(ok=True,selection=selection,scope=scope,
                               counts=dict(nodes=len(candidate.nodes),pipes=len(candidate.pipes),heads=len(candidate.tables.nozzles)))
        except (ValueError,OSError,KeyError) as exc:
            return _fail(str(exc),409)


def install_guards(app) -> None:
    """Replay saved edits before views/exports; block destructive re-extraction."""
    from routes.module_f.jobs import _sess
    inspect_paths = {'/api/module-f/design/preview','/api/module-f/design/emit',
                     '/api/module-f/convert/run','/api/module-f/merge/build',
                     '/api/module-f/merge/preview','/api/module-f/merge/emit'}
    rebuild_paths = {'/api/module-f/design/build','/api/module-f/auto/run'}

    def wrap(fn):
        @functools.wraps(fn)
        def guarded(*args,**kwargs):
            body = request.get_json(silent=True) or {} if request.method=='POST' else request.args
            sid = body.get('sid')
            if not sid:
                return fn(*args,**kwargs)
            try:
                sess = _sess(sid)
            except ValueError:
                return fn(*args,**kwargs)
            with ne.EDIT_LOCK:
                try:
                    if _job_running(sess):
                        return fn(*args,**kwargs)
                    # Explicit rebuilds reconcile history only AFTER successful
                    # extraction. Old history must not lock out table confirmation.
                    if request.path in rebuild_paths:
                        return fn(*args,**kwargs)
                    design = ne.plan_state(sess).get('design')
                    if design:
                        editor = ne.ensure(sess,'design')
                        if editor['conflict']:
                            raise EditError(editor['conflict'])
                    if request.path.endswith(('/merge/preview','/merge/emit')) and (sess.get('merged') or {}).get('combined'):
                        editor = ne.ensure(sess,'merge')
                        if editor['conflict']:
                            raise EditError(editor['conflict'])
                except (ValueError,OSError) as exc:
                    return _fail(str(exc),409)
                return fn(*args,**kwargs)
        return guarded

    for rule in app.url_map.iter_rules():
        if rule.rule in inspect_paths | rebuild_paths:
            app.view_functions[rule.endpoint] = wrap(app.view_functions[rule.endpoint])

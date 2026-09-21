"""Shared live graph editing endpoints for calculation and merged networks."""
from __future__ import annotations

from copy import deepcopy
import functools
import hashlib
import json

from flask import jsonify, request

from core.network_editor import EditError, apply_edit
from core.network_edit_replay import NODE_OPS, rebase_commands
from routes.module_f import network_edit as ne
from routes.module_f.common import _fail
from routes.module_f.jobs import _job_running, route_session


def _target_scope(sess: dict, scope: str, command: dict) -> tuple[str, dict, int]:
    """Plan objects in the merged view edit the original plan, then re-merge."""
    if scope != 'merge':
        return scope,command,0
    from routes.module_f.merge import label_offset_for
    obj = sess['merged']
    nodes = set(map(str,(obj.get('parts') or {}).get('plan',())))
    target = str(command.get('target'))
    kind = command.get('op')
    is_node = kind in NODE_OPS
    row = next((r for r in obj['combined'].pipes if str(r['label'])==target),None)
    in_plan = target in nodes if is_node else bool(row and str(row['in']) in nodes and str(row['out']) in nodes)
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
    elif target not in original_net.pipes:
        return scope,command,0
    for field in (('end',) if kind=='connect' else ('source',) if kind=='paste' else ()):
        label=str(command.get(field,''))
        if not label.isdigit() or str(int(label)-offset) not in original_net.nodes:
            return scope,command,0
    command = deepcopy(command)
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
                cursor: int, *, branch: bool = False) -> None:
    """Persist and accept together. Protected from cancellation mid-commit."""
    fresh = dict(editor,current=candidate,commands=commands,cursor=cursor,conflict=None)
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
                other['revision'] = _revision(other)
                other_state[other_key] = other
        obj.clear()
        obj.update(deepcopy(editor['base_object']))
        ne.sync_object(sess,scope,obj,candidate)
        if scope == 'design':
            _rebuild_merge(sess)
        fresh['revision'] = _revision(fresh)
        pending = {scope:fresh}
        if (sess.get('merged') or {}).get('combined'):
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
                return jsonify(ok=True,scope=scope,catalog=ne.catalog(),**result)
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
                    command = _riser_command(sess,scope,command)
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
                    candidate = deepcopy(editor['base'])
                    for cmd in commands[:cursor]:
                        candidate,selection = apply_edit(candidate,cmd,ne.catalog())
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
                        if selection['kind']=='node':
                            selection = dict(selection,label=str(int(selection['label'])+offset))
                        selection['counts'] = dict(nodes=len(view['nodes']),pipes=len(view['pipes']),
                            heads=len(clone['merged']['combined'].nozzles))
                    return jsonify(ok=True,preview=view,selection=selection,
                                   counts=selection['counts'],revision=shown['revision'],effective_scope=scope)
                from routes.module_f.cancellation import checkpoint
                checkpoint()
                commit_edit(sess,scope,editor,candidate,commands,cursor,branch=action=='apply')
                if selection and offset and selection['kind']=='node':
                    selection = dict(selection,label=str(int(selection['label'])+offset))
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

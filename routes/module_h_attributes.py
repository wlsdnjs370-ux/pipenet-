"""H-only drawing evidence and transactional bulk property corrections.

The existing stable override keys remain the storage contract. Nothing here
changes connectivity, coordinates, lengths, or the operating-head selection.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict
from datetime import datetime, timezone
import hashlib
import json
import math
from pathlib import Path
import threading
import uuid

from flask import Flask, jsonify
from routes.module_f import overrides as ov
from routes.module_f.common import _fail
from routes.module_f.jobs import route_session
from src.pipenet_converter.graph.diameter_definition import POLICY, graph_revision, ordered_path, pattern_preview

_LOCK = threading.RLock()


def apply_head_properties(tables, keys: dict, rows: list) -> None:
    """Carry explicit H head corrections into the authoritative nozzle table."""
    heads={str(label):tuple(key) for label,key in keys.get('node',{}).items() if key[0]=='head'}
    for nozzle in tables.nozzles:
        key=heads.get(str(nozzle['in']))
        for row in rows:
            if tuple(row.get('key',[]))==key and row.get('field') in ('k_factor_si','required_pressure_bar'):
                nozzle[row['field']]=float(row['new'])
                nozzle['definition_policy']=POLICY


def _identity(key: list) -> str:
    return json.dumps(key, ensure_ascii=False, separators=(',', ':'))


def _effective_rows(rows: list) -> dict:
    """Legacy rebuilds enrich only `old`; that is not a concurrent user edit."""
    return {(str(r.get('key')),r.get('field')):{k:v for k,v in r.items() if k!='old'} for r in rows}


def _disk(sess: dict) -> tuple[Path, str]:
    path = Path(ov.path_for(sess.get('key') or 'design'))
    return path, hashlib.sha256(path.read_bytes() if path.is_file() else b'').hexdigest()


def _basis(sess: dict) -> tuple:
    if (sess.get('design_settings') or {}).get('diameter_policy') != POLICY:
        raise ValueError('H 속성 정의를 먼저 실행하세요.')
    board = getattr(sess.get('edit'), 'board', None)
    context = sess.get('_h_diameter_context')
    if board is None or context is None or not sess.get('design'):
        raise ValueError('영역 배관망을 추출한 뒤 속성 정의를 실행하세요.')
    return board, context


def state(sess: dict) -> dict:
    """Return source coordinates, stable identities and explicit evidence for review."""
    from routes.module_f.network_edit import catalog
    board, context = _basis(sess)
    design, records, covered = sess['design'], [], set()
    got, tbl = design['got'], design['tables']
    keys = ov.label_keys(got, board, tbl)
    overrides = ov.ensure_loaded(sess)
    manual = {(tuple(r.get('key', [])), r.get('field')):r for r in overrides}
    catalog_rows = catalog()['pipes']
    labels = {str(v):str(k) for k,v in tbl.pipe_labels.items()}
    for p in tbl.pipes:
        evidence = deepcopy(p.get('bore_provenance') or {})
        key = keys['pipe'].get(str(p['label']))
        members = [r.get('bore_provenance',{}).get('stable_key') for r in evidence.get('continuous_sources',[])]
        members = [list(k) for k in members if k]
        if not key and members:
            key=['run',*sorted(_identity(k) for k in members)]
        if not key:
            # Canonical split/added pipes have no original edge key. They remain
            # editable through the revision-checked network editor, not overrides.
            if not design.get('editor_canonical'):
                continue
            key = ['editor_pipe', str(p['label'])]
        target_keys=members if key[0]=='run' else [key]
        pid = labels.get(str(p['label']))
        ref = got.get('edge_ref', {}).get(pid)
        edges = context.flow.between(*ref) if ref is not None else []
        if key[0]=='run':
            edges=sorted({e for k in target_keys if k[0]=='pipe' for e in context.flow.between(*k[1:])})
        covered.update(edges)
        xy = [[list(board.pts[a][:2]), list(board.pts[b][:2])] for a,b in edges]
        if not xy and key[0]=='vert':
            anchor = board.disks[key[2]][:2] if key[1]=='head' else board.pts[key[2]][:2]
            xy = [[list(anchor),list(anchor)]]
        edit = manual.get((tuple(key),'dia'))
        # Canonical edits already include the older stable-key overrides. Do not
        # let that earlier value mask a later source-text reference transaction.
        pending_edit = edit if not design.get('editor_canonical') or design.get('h_attributes_dirty') else None
        nominal = pending_edit['new'] if pending_edit else p.get('dia') or None
        source = 'user' if pending_edit or evidence.get('manual') else evidence.get('definition_source') or p.get('dia_src') or 'unresolved'
        records.append(dict(id=_identity(key), key=key, keys=target_keys,label=str(p['label']), kind='pipe',
            xy=xy, edges=[list(e) for e in edges], nominal_mm=nominal,
            inner_mm=next((r['inner_mm'] for r in catalog_rows if r['type']==p['type'] and r['dia']==nominal),None),
            schedule=p['type'], source=source, evidence=evidence,
            protected=bool(edit or evidence.get('manual') or evidence.get('text_mm') or
                           any((r.get('bore_provenance') or {}).get('text_mm') for r in evidence.get('continuous_sources',[]))),
            generated=not bool(edges), editable=True))
    # Other retained parts of the drawing can serve as a representative without
    # adding their heads to the calculation scenario.
    for edge, evidence in context.decisions.items():
        if edge in covered:
            continue
        key=['reference',*edge]
        records.append(dict(id=_identity(key),key=key,label=f'{edge[0]}–{edge[1]}',kind='reference',
            xy=[[list(board.pts[n][:2]) for n in edge]],edges=[list(edge)],
            nominal_mm=evidence.get('text_mm') if not evidence.get('block_export') else None,
            source=evidence.get('definition_source','unresolved'),evidence=deepcopy(evidence),editable=False))
    active = set((sess.get('worst') or {}).get('heads') or [])
    for index in sorted(active):
        key=['head',index]
        records.append(dict(id=_identity(key),key=key,kind='head',label=f'H{index}',
            node_label=next((str(label) for label,k in keys['node'].items() if list(k)==key),None),
            xy=[list(board.disks[index][:2])],editable=True,
            properties={f:r['new'] for (k,f),r in manual.items() if k==tuple(key)}))
    _, disk_hash = _disk(sess)
    sess.setdefault('_h_attribute_disk_hash',disk_hash)
    revision = graph_revision(board.pts,board.edges,{i:str(k) for i,k in enumerate(board.disk_kinds)},
                              board.sources, [overrides,sess.get('design_settings'),disk_hash,records])
    return dict(ok=True,revision=revision,records=records,annotations=[asdict(a) for a in sess.get('_h_diameter_annotations',[])],
                summary=context.summary,catalog=catalog_rows,
                pending=sum(not r.get('nominal_mm') for r in records if r['kind']=='pipe'),
                can_undo=bool(sess.get('_h_attribute_undo')),policy=POLICY)


def _checked(sess: dict, body: dict) -> tuple[dict, dict]:
    if (sess.get('job') or {}).get('state')=='run':
        raise ValueError('계산이 끝난 뒤 수정하세요.')
    current=state(sess)
    if body.get('revision')!=current['revision']:
        raise ValueError('도면 또는 속성이 변경되었습니다. 최신 결과를 다시 확인하세요.')
    path,disk_hash = _disk(sess)
    raw=json.loads(path.read_bytes()) if path.is_file() else {}
    disk_rows=(raw.get('items') or []) if isinstance(raw,dict) else raw
    local=_effective_rows(ov.ensure_loaded(sess))
    disk=_effective_rows(disk_rows)
    if any(local.get(k)!=r for k,r in disk.items()) or (
            disk_hash!=sess.get('_h_attribute_disk_hash') and disk!=local):
        raise ValueError('다른 창에서 저장한 수정이 있습니다. 도면을 다시 열어 최신 속성을 불러오세요.')
    return current,{r['id']:r for r in current['records']}


def _selected(body: dict, records: dict, field: str='ids') -> list[dict]:
    ids=body.get(field)
    if not isinstance(ids,list) or not ids or len(ids)>5000 or len(ids)!=len(set(ids)):
        raise ValueError('대상을 1~5000개 선택하세요.')
    if any(i not in records for i in ids):
        raise ValueError('현재 도면에 없는 대상입니다.')
    return [records[i] for i in ids]


def _write(sess: dict, rows: list) -> None:
    """Atomic write-through, retaining unrelated ops and cancellation rollback."""
    path,_ = _disk(sess)
    previous=path.read_bytes() if path.is_file() else None
    raw=json.loads(previous) if previous else {'version':1}
    if not isinstance(raw,dict):
        raw={'version':1,'items':raw}
    raw['items']=rows
    path.parent.mkdir(parents=True,exist_ok=True)
    tmp=path.with_suffix('.'+uuid.uuid4().hex+'.tmp')
    tmp.write_text(json.dumps(raw,ensure_ascii=False,indent=2),encoding='utf-8')
    tmp.replace(path)
    written=path.read_bytes()
    sess['_h_attribute_disk_hash']=hashlib.sha256(written).hexdigest()
    from routes.module_f.cancellation import current_operation
    operation=current_operation()
    if operation is not None:
        def rollback():
            with _LOCK:
                if path.is_file() and path.read_bytes()==written:
                    if previous is None:
                        path.unlink()
                    else:
                        tmp.write_bytes(previous)
                        tmp.replace(path)
        operation.rollback_callbacks.append(rollback)
    ov.save(sess,rows)
    sess['design']['h_attributes_dirty']=True


def register(app: Flask) -> None:
    """Reuse the shared per-session cancellation/serialization guard, H opt-in only."""
    @app.get('/api/module-f/h-attributes/state')
    @route_session()
    def h_attribute_state(sess,body):
        try:
            with _LOCK: return jsonify(state(sess))
        except ValueError as exc: return _fail(str(exc),409)

    @app.post('/api/module-f/h-attributes/copy-preview')
    @route_session(post=True)
    def h_attribute_copy_preview(sess,body):
        try:
            with _LOCK:
                current,records=_checked(sess,body)
                source=_selected(body,records,'source_ids')
                target=_selected(body,records)
                if any(not r.get('edges') for r in source+target):
                    raise ValueError('평면의 연속 배관 경로를 선택하세요. 헤드·수직 접속은 일괄 수정으로 지정합니다.')
                points=_basis(sess)[0].pts
                by_edge={tuple(e):r for r in source for e in r['edges']}
                _,path=ordered_path(by_edge,body.get('source_anchor'))
                profile=[dict(length_mm=math.dist(points[a][:2],points[b][:2]),nominal_mm=by_edge[a,b]['nominal_mm'],
                              source_id=by_edge[a,b]['id']) for a,b in path]
                raw=pattern_preview(points,profile,[e for r in target for e in r['edges']],
                                    anchor=body.get('target_anchor'),reverse=body.get('reverse') is True)
                proposed={tuple(r['edge']):r for r in raw}
                preview=[]
                for r in target:
                    parts=[proposed[tuple(e)] for e in r['edges']]
                    candidates=sorted({v for part in parts for v in part['candidates']})
                    preview.append(dict(id=r['id'],label=r['label'],current_mm=r['nominal_mm'],
                        nominal_mm=parts[0]['nominal_mm'],needs_choice=len(candidates)>1,
                        candidates=candidates,protected=r.get('protected',False)))
                sess['_h_copy_preview']={'revision':current['revision'],'rows':preview,'source_ids':body['source_ids']}
                return jsonify(ok=True,rows=preview,revision=current['revision'])
        except (ValueError,TypeError,KeyError) as exc: return _fail(str(exc),409)

    @app.post('/api/module-f/h-attributes/apply')
    @route_session(post=True)
    def h_attribute_apply(sess,body):
        try:
            with _LOCK:
                current,records=_checked(sess,body)
                targets=_selected(body,records)
                rows=deepcopy(ov.ensure_loaded(sess));before=deepcopy(rows)
                action=body.get('action','manual')
                if action not in ('manual','annotation','paste'): raise ValueError('알 수 없는 수정 방식입니다.')
                annotation=None
                if action=='annotation':
                    annotation=next((a for a in current['annotations'] if a['identity']==body.get('annotation_id')),None)
                    if annotation is None: raise ValueError('도면 표기를 선택하세요.')
                pasted={}
                if action=='paste':
                    cached=sess.get('_h_copy_preview') or {}
                    if cached.get('revision')!=current['revision']: raise ValueError('붙여넣기 미리보기를 다시 실행하세요.')
                    pasted={r['id']:r for r in cached['rows']}
                    if set(pasted)!=set(body['ids']): raise ValueError('미리보기와 대상이 다릅니다.')
                skipped=0
                for target in targets:
                    if not target['editable']: raise ValueError('추출 영역 밖 참조 배관은 복사 원본으로만 사용할 수 있습니다.')
                    key=target['key'];kind=key[0]
                    if kind=='head':
                        if action!='manual': raise ValueError('헤드는 일괄 속성 수정으로 지정하세요.')
                        fields={k:v for k,v in (body.get('head_properties') or {}).items() if v is not None}
                        if not fields or set(fields)-{'k_factor_si','required_pressure_bar'}:
                            raise ValueError('헤드 K 값 또는 최소 압력을 입력하세요.')
                    else:
                        if action!='manual' and target.get('protected') and body.get('overwrite') is not True:
                            skipped+=1;continue
                        raw=annotation['nominal_mm'] if annotation else body.get('nominal_mm')
                        if action=='paste':
                            candidate=pasted[target['id']]
                            choice=(body.get('choices') or {}).get(target['id'])
                            if candidate['needs_choice'] and choice is None: raise ValueError('관경 경계를 가로지르는 구간의 관경을 직접 선택하세요.')
                            raw=choice if choice is not None else candidate['nominal_mm']
                        value=float(raw)
                        if not math.isfinite(value) or not value.is_integer(): raise ValueError('관경은 규격표의 호칭경을 선택하세요.')
                        if not any(r['type']==target['schedule'] and r['dia']==value for r in current['catalog']):
                            raise ValueError(f"{target['label']}: 해당 재질에 없는 관경입니다.")
                        fields={'dia':int(value)}
                    for field,value in fields.items():
                        reason=str(body.get('reason') or {'manual':'선택 일괄 수정','annotation':'도면 표기 수동 연결','paste':'대표 관경 패턴 수동 대응'}[action])[:200]
                        for member_key in target.get('keys',[key]):
                            member_kind=member_key[0]
                            parsed,why=(int(value),None) if member_kind=='vert' and field=='dia' else ov.validate(member_kind,field,value)
                            if why: raise ValueError(why)
                            rows=ov.put(rows,member_key,member_kind,field,parsed,old=target.get('nominal_mm'),reason=reason,
                                        at=datetime.now(timezone.utc).isoformat())
                            rows[-1]['definition_origin']={'action':action,'annotation':annotation,
                                'source_ids':(sess.get('_h_copy_preview') or {}).get('source_ids') if action=='paste' else None}
                if rows==before: raise ValueError('적용 대상이 없습니다. 기존 도면·사용자 관경 보호 여부를 확인하세요.')
                if len(rows)>5000: raise ValueError('저장 가능한 수정 개수를 초과했습니다.')
                _write(sess,rows)
                sess.setdefault('_h_attribute_undo',[]).append({'before':before,'after':deepcopy(rows)})
                sess['_h_attribute_undo']=sess['_h_attribute_undo'][-30:]
                return jsonify(ok=True,skipped=skipped,needs_rebuild=True)
        except (ValueError,TypeError,KeyError,OSError) as exc: return _fail(str(exc),409)

    @app.post('/api/module-f/h-attributes/undo')
    @route_session(post=True)
    def h_attribute_undo(sess,body):
        try:
            with _LOCK:
                _checked(sess,body)
                stack=sess.get('_h_attribute_undo') or []
                if not stack: raise ValueError('되돌릴 일괄 수정이 없습니다.')
                item=stack[-1]
                if _effective_rows(ov.ensure_loaded(sess))!=_effective_rows(item['after']): raise ValueError('다른 편집이 추가되어 일괄 되돌릴 수 없습니다. 편집 기록에서 확인하세요.')
                _write(sess,item['before']);stack.pop()
                return jsonify(ok=True,needs_rebuild=True)
        except (ValueError,OSError) as exc: return _fail(str(exc),409)

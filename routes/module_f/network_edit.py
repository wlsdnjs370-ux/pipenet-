"""Shared confirmed-network editor: library, history, table/KFP adapters.

Histories are tied to an exact base graph fingerprint, never to a reused label.
JSON commands are saved atomically and replayed after reopening the same graph.
"""
from __future__ import annotations

from copy import deepcopy
from functools import lru_cache
import hashlib
import json
from pathlib import Path
import threading
from datetime import datetime, timezone
from typing import Any

from core.network_editor import EditError, Network, apply_edit, merge_reason, terminal_pipe
from routes.module_f.common import _boot

EDIT_LOCK = threading.RLock()
HISTORY_DIR = Path(__file__).resolve().parents[2] / 'data' / 'network_editor'


@lru_cache(maxsize=1)
def catalog() -> dict:
    """Read actual SLF bore/nozzle data and the Solver fitting library."""
    _boot()
    from kfp_sdf_converter import _load_slf_authoritative, _SLF_SCHEDULE_CFACTOR
    from services.cad_import.design.sdf_post import SCHEDULE_DEFS
    from services.cad_import.design.fitting import FITTING_LIB_ID
    lib = _load_slf_authoritative()
    allowed = {name: {round(d*1000) for d, _ in sizes} for name, _, sizes in SCHEDULE_DEFS}
    pipes = []
    for name, value in lib['schedules'].items():
        for dn, inner in value['sizes'].items():
            if dn in allowed.get(name, set()):
                pipes.append(dict(type=name, dia=dn, inner_mm=inner,
                                  c=_SLF_SCHEDULE_CFACTOR.get(name,120),
                                  roughness_mm=value['roughness_mm']))
    if not pipes:
        raise EditError('PIPENET SLF 관경 라이브러리를 읽지 못했습니다.')
    root = Path(__file__).resolve().parents[2]
    source = root/'cad_project_editor_g'/'fittings_library_v3.json'
    raw = json.loads(source.read_text(encoding='utf-8'))
    fittings = [dict(id=r['id'], name=r.get('name_ko') or r['id'],
                    category=r.get('category'),
                    lengths={str(d):v.get('value') for d,v in r.get('equivalent_lengths',{}).items()})
                for r in raw['items']]
    nozzles = [dict(id=k, name=v['desc'], k_factor_si=v['k_si']*60000*(100000**.5),
                    min_bar=v['min_p_pa']/100000, max_bar=v['max_p_pa']/100000)
               for k,v in lib['nozzles'].items() if v['k_si'] > 0]
    return dict(pipes=pipes,fittings=fittings,nozzles=nozzles,
                fitting_aliases={k:v for k,v in FITTING_LIB_ID.items()
                                 if k not in ("tee-run", "cross-run", "cross")},
                sources=[Path(lib['_path']).name, source.name])


def fingerprint(network: Network) -> str:
    """All hydraulic values and identities belong to the revision."""
    body = dict(nodes={k:dict(xyz=n.xyz,row=n.row) for k,n in network.nodes.items()},
                pipes={k:dict(a=p.a,b=p.b,row=p.row) for k,p in network.pipes.items()},
                nozzles=network.tables.nozzles,fittings=network.tables.fittings,
                equipment=network.tables.equipment,
                pumps=getattr(network.tables,'pumps',[]), valves=getattr(network.tables,'valves',[]))
    return hashlib.sha256(json.dumps(body,sort_keys=True,ensure_ascii=False,
                                     allow_nan=False).encode()).hexdigest()


def plan_state(sess: dict) -> dict:
    """Read the plan even while a different drawing slot is active."""
    from routes.module_f.slots import _slot_active
    return sess if _slot_active(sess)=='plan' else (sess.get('slots') or {}).get('plan',{})


def _container(sess: dict, scope: str) -> tuple[dict, str, dict]:
    if scope == 'merge':
        got = sess.get('merged')
        if not got or got.get('combined') is None:
            raise EditError('먼저 배관망을 결합하세요.')
        return sess, 'merge_editor', got
    state = plan_state(sess)
    design = state.get('design')
    if not design:
        raise EditError('먼저 수리계산 입력표를 확정하세요.')
    return state, 'network_editor', design


def _network(sess: dict, scope: str, obj: dict) -> Network:
    if scope == 'merge':
        parts = obj.get('parts') or {}
        protected = set()
        for a in ('system','machineroom'):
            protected |= set(map(str,parts.get(a,()))) & set(map(str,parts.get('plan',())))
        return Network.from_tables(obj['combined'],protected=protected)
    from routes.module_f.api_design import _label_keys
    state = plan_state(sess)
    lk = _label_keys(state,obj['got'],obj['tables'])
    meta = (obj['got'].get('kfp') or {}).get('nodes_meta_runtime',{})
    pos = {lab:tuple(meta[nid]['coords']) for lab,nid in lk['nid'].items()
           if nid in meta and len(meta[nid].get('coords',[]))==3}
    return Network.from_tables(obj['tables'],positions=pos)


def _path(sess: dict, scope: str) -> Path:
    state = plan_state(sess)
    # Drawing keys can contain user text. Only the digest enters a filename.
    identity = str(state.get('key') or sess['id']) + ':' + scope
    digest = hashlib.sha256(identity.encode()).hexdigest()
    return HISTORY_DIR/(digest+'.json')


def save_history(sess: dict, scope: str, editor: dict) -> None:
    """Persist before committing in memory; never expose a half-written file."""
    path = _path(sess,scope)
    path.parent.mkdir(parents=True,exist_ok=True)
    previous = path.read_bytes() if path.is_file() else None
    if path.is_file() and editor.get('disk_hash') != hashlib.sha256(path.read_bytes()).hexdigest():
        raise EditError('다른 창에서 같은 도면의 편집을 저장했습니다. 도면을 다시 열어 최신 기록을 확인하세요.')
    body = dict(version=1,base=editor['base_hash'],commands=editor['commands'],
                library=hashlib.sha256(json.dumps(catalog(),sort_keys=True).encode()).hexdigest(),
                cursor=editor['cursor'],drawing=plan_state(sess).get('key'),scope=scope,
                saved_at=datetime.now(timezone.utc).isoformat())
    tmp = path.with_suffix('.tmp')
    tmp.write_text(json.dumps(body,ensure_ascii=False,indent=2),encoding='utf-8')
    tmp.replace(path)
    editor['disk_hash'] = hashlib.sha256(path.read_bytes()).hexdigest()
    # A stop can arrive during the final atomic write. The cancellation layer
    # restores its session snapshot; restore this matching disk revision too.
    from routes.module_f.cancellation import current_operation
    operation = current_operation()
    if operation is not None:
        written = editor['disk_hash']
        def write_rollback():
            with EDIT_LOCK:
                if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != written:
                    return  # A newer, independently accepted edit owns the file.
                if previous is None:
                    path.unlink()
                else:
                    tmp.write_bytes(previous)
                    tmp.replace(path)
        operation.rollback_callbacks.append(write_rollback)


def _history_versions(sess: dict, scope: str) -> list[tuple[Path, dict]]:
    """Find preserved versions belonging to this drawing and scope only."""
    directory = HISTORY_DIR / 'archive' / _path(sess,scope).stem
    if not directory.is_dir():
        return []
    return [(p,json.loads(p.read_text(encoding='utf-8')))
            for p in sorted(directory.glob('*.json'),key=lambda p:p.stat().st_mtime_ns,reverse=True)]


def backup_history(sess: dict, scope: str, disk_hash: str | None) -> None:
    """Keep exact saved bytes before switching bases; never erase old commands."""
    path = _path(sess,scope)
    if not path.is_file():
        return
    raw = path.read_bytes()
    digest = hashlib.sha256(raw).hexdigest()
    if disk_hash != digest:
        raise EditError('다른 창에서 편집 기록이 바뀌었습니다. 최신 기록을 다시 읽어 주세요.')
    directory = HISTORY_DIR / 'archive' / path.stem
    directory.mkdir(parents=True,exist_ok=True)
    archive = directory / (digest+'.json')
    if not archive.exists():
        tmp = archive.with_suffix('.tmp')
        tmp.write_bytes(raw)
        tmp.replace(archive)


def ensure(sess: dict, scope: str, *, rebuilt: bool = False) -> dict:
    """Load/replay edits once for each freshly constructed base graph."""
    state,key,obj = _container(sess,scope)
    existing = state.get(key)
    if existing and existing['object'] is obj and not rebuilt:
        return existing
    base = _network(sess,scope,obj)
    base_hash = fingerprint(base)
    saved = None
    if existing:
        # A rejected history still belongs to its ORIGINAL base, not the new
        # graph shown in the conflict state. Keep that identity when rebuilding.
        saved = existing.get('saved_history') if existing.get('conflict') else None
        if saved is None:
            saved = dict(base=existing['base_hash'],commands=existing['commands'],cursor=existing['cursor'],
                         library=existing.get('library'))
    elif _path(sess,scope).is_file():
        saved = json.loads(_path(sess,scope).read_text(encoding='utf-8'))
    lib_hash = hashlib.sha256(json.dumps(catalog(),sort_keys=True).encode()).hexdigest()
    disk_hash = (hashlib.sha256(_path(sess,scope).read_bytes()).hexdigest()
                 if _path(sess,scope).is_file() else None)
    changed = bool(saved and (saved['base'] != base_hash or saved.get('library',lib_hash) != lib_hash))
    notice = None
    if rebuilt and changed:
        # Explicit table confirmation starts a new calculation basis. Identical
        # histories are replayed; unrelated labels are never guessed or applied.
        backup_history(sess,scope,existing.get('disk_hash') if existing else disk_hash)
        count = (saved or {}).get('cursor',0)
        saved = next((body for _,body in _history_versions(sess,scope)
                      if body['base']==base_hash and body.get('library',lib_hash)==lib_hash),None)
        notice = ('같은 기준 배관망의 편집 기록을 복원했습니다.' if saved else
                  f'새 기준 배관망으로 표를 확정했습니다. 이전 편집 {count}건은 편집 기록에 별도 보관했습니다.')
    commands = (saved or {}).get('commands',[])
    cursor = (saved or {}).get('cursor',0)
    if saved and saved['base'] != base_hash and not cursor:
        commands = []  # Redo commands also belong to their original base graph.
    editor = dict(object=obj,base=base,base_hash=base_hash,current=base,library=lib_hash,
                  disk_hash=disk_hash,saved_history=deepcopy(saved),notice=notice,
                  base_object=deepcopy(obj),
                  commands=commands,cursor=cursor,revision=base_hash,conflict=None)
    if cursor and (saved['base'] != base_hash or saved.get('library',lib_hash) != lib_hash):
        editor['conflict'] = '기준 배관망이 바뀌어 저장된 편집을 적용하지 않았습니다. 기존 편집 기록을 확인하고 편집 초기화를 선택하세요.'
    else:
        for command in commands[:cursor]:
            base,_ = apply_edit(base,command,catalog())
        editor['current'] = base
        editor['revision'] = fingerprint(base)
        if cursor:
            sync_object(sess,scope,obj,base)
    if rebuilt and changed:
        save_history(sess,scope,editor)
    state[key] = editor
    return editor


def accept_rebuilt(sess: dict, scope: str = 'design') -> dict:
    """Attach the appropriate history after a successful, explicit table build."""
    with EDIT_LOCK:
        editor = ensure(sess,scope,rebuilt=True)
        if editor.get('notice'):
            print('[편집 기록] '+editor['notice'])
        return editor


def archived_histories(sess: dict, scope: str) -> list[dict]:
    """Read-only recovery information shown in the existing history dialog."""
    return [dict(file=str(path.relative_to(HISTORY_DIR)),base=body['base'],
                 cursor=body.get('cursor',0),saved_at=body.get('saved_at'),
                 commands=body.get('commands',[])) for path,body in _history_versions(sess,scope)]


def sync_object(sess: dict, scope: str, obj: dict, net: Network) -> None:
    """Accept the edited graph into the exact models used by every exporter."""
    tables = deepcopy(net.to_tables())
    if scope == 'merge':
        from routes.module_f.merge import adopt_late_nodes, combined_summary
        known = {str(n['label']) for n in obj['combined'].nodes}
        obj['combined'] = tables
        adopt_late_nodes(tables,obj.setdefault('parts',{}),known)
        for kind, labels in obj['parts'].items():
            obj['parts'][kind] = [x for x in labels if str(x) in net.nodes]
        sess['merge_summary'] = combined_summary(obj)
        sess.pop('merge_files',None)
        return
    state = plan_state(sess)
    from routes.module_f.api_design import _label_keys
    keys = deepcopy(_label_keys(state,obj['got'],obj['tables']))
    got = deepcopy(obj['got'])
    kfp = got.setdefault('kfp',{})
    from kfp_sdf_converter import _slf_pipe_library_for_kfp, _SLF_TO_KSOLVER_STD
    kfp.setdefault('version','4.0-NFPA13-EQL')
    kfp['pipe_library'] = _slf_pipe_library_for_kfp()
    nodes = kfp.setdefault('nodes_meta_runtime',{})
    pipes = kfp.setdefault('pipe_data',{})
    nid_of = keys.setdefault('nid',{})
    pid_of = {str(lab):str(pid) for pid,lab in getattr(obj['tables'],'pipe_labels',{}).items()}
    # Recover mappings from pipe endpoints when old auto tables have none.
    # Resolve identities from the ORIGINAL endpoints, before a split has replaced
    # one endpoint. Using net here aliases a new split node to the old end node.
    for row in obj['tables'].pipes:
        lab = str(row['label'])
        rec = pipes.get(pid_of.get(lab,lab),{})
        for n,field in ((str(row['in']),'start'),(str(row['out']),'end')):
            if rec.get(field) in nodes:
                nid_of.setdefault(n,rec[field])
    used = set(nodes)
    nozzle_of = {str(r['in']):r for r in tables.nozzles}
    for lab,n in net.nodes.items():
        if lab not in nid_of:
            name = 'NE'+lab
            while name in used:
                name += '_'
            nid_of[lab] = name
            used.add(name)
        nid = nid_of[lab]
        rec = deepcopy(nodes.get(nid) or {'type_id':'base','type':'기본','category_id':''})
        rec.update(id=nid,coords=list(n.xyz),elevation_m=n.xyz[2])
        head = n.row.get('editor_head','unchanged')
        if nid not in nodes and lab in nozzle_of and head=='unchanged':
            nz = nozzle_of[lab]
            spec = next((r for r in catalog()['nozzles'] if r['id']==nz.get('lib')),None)
            if spec:
                head = dict(spec,flow_lmin=float(nz.get('flow_lmin') or float(nz.get('flow_m3s') or 0)*60000))
        if head != 'unchanged':
            rec.update(type_id='head' if head else 'base',type='헤드' if head else '기본')
            if head:
                rec.update(k_factor_si=head['k_factor_si'],head_spec_name=head['id'],
                           required_pressure_bar=(head['flow_lmin']/head['k_factor_si'])**2)
            else:
                rec.update(k_factor_si=None,head_spec_name=None,required_pressure_bar=0)
        nodes[nid] = rec
    kfp['nodes_meta_runtime'] = {nid_of[lab]:nodes[nid_of[lab]] for lab in net.nodes}
    new_pipes, labels = {}, {}
    for lab,p in net.pipes.items():
        pid = pid_of.get(lab,lab)
        if lab not in pid_of:
            while pid in pipes or pid in new_pipes:
                pid = 'E'+pid
        rec = deepcopy(pipes.get(pid) or {})
        rec.update(start=nid_of[p.a],end=nid_of[p.b],length_m=float(p.row['length']),
                   nominal_mm=p.row.get('dia'),C=p.row.get('c',120))
        # Explicit editor library choice overrides legacy default-standard sizing.
        chosen = p.row if p.row.get('inner_mm') is not None else None
        if chosen is None and pid not in pipes:
            chosen = next((r for r in catalog()['pipes'] if r['type']==p.row.get('type') and r['dia']==p.row.get('dia')),None)
            if chosen is None:
                raise EditError(f'{lab}: 내경을 찾을 수 없습니다. 재질·관경을 먼저 지정하세요.')
        if chosen:
            rec.update(diameter=chosen['inner_mm'],pipe_type=chosen['type'],
                       type=_SLF_TO_KSOLVER_STD.get(chosen['type'],chosen['type']),
                       roughness_mm=chosen.get('roughness_mm',0))
        for field in ('flow','flow_rate','pressure','velocity','pressure_loss',
                      'flow_lpm','velocity_mps','headloss_m'):
            rec.pop(field,None)
        new_pipes[pid] = rec
        labels[pid] = lab
    kfp['pipe_data'] = new_pipes
    tables.pipe_labels = labels
    keys['nid'] = {lab:nid for lab,nid in nid_of.items() if lab in net.nodes}
    keys['node'] = {lab:k for lab,k in keys.get('node',{}).items() if lab in net.nodes}
    keys['pipe'] = {lab:k for lab,k in keys.get('pipe',{}).items() if lab in net.pipes}
    obj.update(tables=tables,got=got,keys=keys,editor_modified=True)
    if state.get('auto'):
        state['auto']['tables'] = tables
    for field in ('design_sdf_path','design_slf_path','design_has_path','worst_kfp_path'):
        state.pop(field,None)


def public_state(editor: dict) -> dict:
    """Small edit state including loss/source information for the inspector."""
    net = editor['current']
    head_specs = {}
    libraries = {r['id']: r for r in catalog()['nozzles']}
    for row in net.tables.nozzles:
        if row.get('lib') in libraries:
            head_specs[str(row['in'])] = dict(libraries[row['lib']],
                flow_lmin=row.get('flow_lmin') or float(row.get('flow_m3s',0))*60000)
    return dict(revision=editor['revision'],undo=editor['cursor'],
                redo=len(editor['commands'])-editor['cursor'],conflict=editor['conflict'],
                history=editor['commands'],
                nodes=[dict(label=k,xyz=n.xyz,head=k in net.heads(),protected=k in net.protected,
                            terminal_pipe=terminal_pipe(net,k),merge_reason=merge_reason(net,k),
                            **({'head_spec':head_specs[k]} if k in head_specs else {}))
                       for k,n in net.nodes.items()],
                pipes=[dict(p.row,**{'in':p.a,'out':p.b}) for p in net.pipes.values()],
                equipment=net.tables.equipment)


def project(sess: dict, scope: str, net: Network) -> dict:
    """The preview uses the same projection as saved SDF and the main canvas."""
    tables = net.to_tables()
    if scope == 'merge':
        from routes.module_f.merge import bake_combined_iso, adopt_late_nodes
        obj = deepcopy(sess['merged'])
        known = {str(n['label']) for n in obj['combined'].nodes}
        obj['combined'] = tables
        adopt_late_nodes(tables,obj.setdefault('parts',{}),known)
        nodes = bake_combined_iso(obj,iso_z_scale=float(sess.get('editor_merge_z',1)))[0] if sess.get('editor_merge_iso') else tables.nodes
    else:
        from services.cad_import.design.emit import display_tables
        from routes.module_f.api_design import _DEFAULT_SETTINGS, _view_opts, _valve_label
        cfg = dict(plan_state(sess).get('design_settings') or _DEFAULT_SETTINGS)
        nodes = display_tables(tables,**_view_opts(cfg),
                     iso_ref_label=_valve_label(tables) if cfg['lift_ref']=='valve' else None)[0].nodes
    return dict(nodes=[dict(label=str(r['label']),x=r['x'],y=r['y']) for r in nodes],
                pipes=[dict(label=k,a=p.a,b=p.b) for k,p in net.pipes.items()])
